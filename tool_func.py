import pymysql
import json
import yaml
from kubernetes import client, config
import urllib3

urllib3.disable_warnings()


class MySQLHandler:
    def __init__(self, conn):
        if isinstance(conn, pymysql.connections.Connection):
            self._conn = conn
            self._cursor = conn.cursor()
        else:
            raise TypeError

    def select(self, sql):
        self._cursor.execute(sql)
        result = self._cursor.fetchall()
        self._cursor.close()
        self._conn.close()
        return result


def get_deploy_info():
    # 集群context
    context_list = []
    # 集群部署信息
    deploy_info = {}
    # 集群总cpu和内存
    cluster_total_resources = {}
    # 集群中各namespace的资源配额
    cluster_namespaces_quota = {}

    with open("static/kubeconfig/config", encoding="utf-8") as kubeconfig:
        kc = yaml.safe_load(kubeconfig)
        for item in kc.get("contexts"):
            context_list.append(item.get("name"))
    for context in context_list:
        deploy_info[context] = {}
        cluster_namespaces_quota[context] = {}
        namespaces_names = []
        config.load_kube_config(config_file="static/kubeconfig/config", context=context)
        namespaces = client.CoreV1Api().list_namespace()
        nodes = client.CoreV1Api().list_node()
        cluster_total_resources[context] = {"cpu": 0, "memory": 0}
        # 通过集群各节点资源计算集群worker节点资源总额（总资源减去master资源）
        for node in nodes.items:
            cluster_total_resources[context]["cpu"] += int(node.status.capacity.get("cpu"))
            cluster_total_resources[context]["memory"] += int(node.status.capacity.get("memory").split("Ki")[0])
        if context == "fcp":
            cluster_total_resources[context]["cpu"] -= 48 * 3
            cluster_total_resources[context]["memory"] = cluster_total_resources[context]["memory"] // (
                    1024 * 1024) - 376 * 3
        else:
            cluster_total_resources[context]["cpu"] -= 56 * 3
            cluster_total_resources[context]["memory"] = cluster_total_resources[context]["memory"] // (
                    1024 * 1024) - 251 * 3
        # 获取集群namespace
        for namespace in namespaces.items:
            namespaces_names.append(namespace.metadata.name)
        # 获取集群中各namespace的应用部署信息（通过deploy或statefulset的labels标签获取应用名称）以及资源配额
        for ns in namespaces_names:
            # 集群namespace应用部署信息
            deploy_info[context][ns] = {}
            # 集群namespace资源配额信息
            cluster_namespaces_quota[context][ns] = {"cpu": 0, "memory": 0}
            # 绿区fcp集群的statefulset的部署信息
            if context == "fcp":
                statefulsets = client.AppsV1Api().list_namespaced_stateful_set(namespace=ns)
                for sts in statefulsets.items:
                    if sts.metadata.labels.get("app"):
                        deploy_info[context][ns][sts.metadata.labels.get("app")] = sts.spec.replicas
            # 集群deployment的部署信息
            deploys = client.ExtensionsV1beta1Api().list_namespaced_deployment(namespace=ns)
            for deploy in deploys.items:
                # 没label的就舍弃
                if deploy.metadata.labels.get("application_name"):
                    deploy_info[context][ns][deploy.metadata.labels.get("application_name")] = deploy.spec.replicas
                elif deploy.metadata.labels.get("app"):
                    deploy_info[context][ns][deploy.metadata.labels.get("app")] = deploy.spec.replicas

            # 集群namespace资源配额
            quotas = client.CoreV1Api().list_namespaced_resource_quota(namespace=ns)
            for quota in quotas.items:
                cpu = quota.spec.hard.get("limits.cpu")
                memory = quota.spec.hard.get("limits.memory")
                cluster_namespaces_quota[context][ns]["cpu"] = int(cpu)
                # 单位Gi
                if quota.spec.hard.get("limits.memory")[-2:] == "Gi":
                    cluster_namespaces_quota[context][ns]["memory"] = int(memory.split("Gi")[0])
                elif quota.spec.hard.get("limits.memory")[-2:] == "Ti":
                    cluster_namespaces_quota[context][ns]["memory"] = int(memory.split("Ti")[0]) * 1024
                elif quota.spec.hard.get("limits.memory")[-2:] == "Mi":
                    cluster_namespaces_quota[context][ns]["memory"] = int(memory.split("Mi")[0]) // 1024
                elif quota.spec.hard.get("limits.memory")[-2:] == "Ki":
                    cluster_namespaces_quota[context][ns]["memory"] = int(memory.split("Ki")[0]) / 1024 // 1024

    return deploy_info, cluster_total_resources, cluster_namespaces_quota


def get_namespace_project():
    with open('config/config.json', encoding='utf-8') as config:
        config = json.load(config)
    conn = pymysql.connect(**(config["mysql"]))
    mh = MySQLHandler(conn)
    sql = "select PROJECT_ID, PROJECT_NAME from dps_pjm_project"
    return mh.select(sql)


def generate_table_data():
    # 从数据库中扩区项目信息（元组）
    # (('01fd132b9a74487cb3a423d5aa74ff2e', '财务公司头寸管理'), ...)
    namespaces_tuple_info = get_namespace_project()
    # 将数据库的project_id和project_name的元组信息转换为字典
    namespaces_info = {}
    for item in namespaces_tuple_info:
        namespaces_info[item[0]] = item[1]

    # 从集群获取应用部署信息，集群资源总量，集群namespace资源分配信息
    deploy_info, cluster_total_resources, cluster_namespaces_quota = get_deploy_info()

    # 用于填充运行情况分析table数据
    cluster_table_data = {}
    for key in deploy_info.keys():
        cluster_table_data[key] = [["项目名称", "服务名称", "实例数"]]
        for ns in deploy_info.get(key).keys():
            postfix = ns.split('-')[-1]
            if postfix == 'fcp' or postfix == 'microservice' or postfix == 'backend':
                for replica_info in deploy_info.get(key).get(ns).items():
                    tmp = []
                    for prj_id in namespaces_info.keys():
                        if prj_id in ns:
                            tmp.append(namespaces_info[prj_id])
                    tmp += list(replica_info)
                    cluster_table_data[key].append(tmp)
        # 合计行
        total = []
        # 统计项目数
        ns_temp = []
        # 统计应用副本数
        replicas_count = 0
        for index in range(1, len(cluster_table_data[key])):
            replicas_count += int(cluster_table_data[key][index][2])
            ns_temp.append(cluster_table_data[key][index][0])
        ns_temp = list(set(ns_temp))
        total.append(str(len(ns_temp)))
        # 统计应用数
        total.append(str(len(cluster_table_data[key]) - 1))
        total.append(str(replicas_count))
        cluster_table_data[key].append(total)

    # 为运行情况分析的扇形图提供数据
    cluster_pie_data = {}
    for cluster in cluster_namespaces_quota.keys():
        # cpu: 核，memory: Gi
        cluster_pie_data[cluster] = {
            "cpu": {"allocated": 0, "total": 0},
            "memory": {"allocated": 0, "total": 0}
        }
        allocated_cpu = 0
        allocated_memory = 0
        for ns in cluster_namespaces_quota.get(cluster).keys():
            allocated_cpu += cluster_namespaces_quota.get(cluster).get(ns).get("cpu")
            allocated_memory += cluster_namespaces_quota.get(cluster).get(ns).get("memory")

        cluster_pie_data[cluster]["cpu"]["allocated"] = allocated_cpu
        cluster_pie_data[cluster]["memory"]["allocated"] = allocated_memory
        cluster_pie_data[cluster]["cpu"]["total"] = cluster_total_resources.get(cluster).get("cpu")
        cluster_pie_data[cluster]["memory"]["total"] = cluster_total_resources.get(cluster).get("memory")

    return cluster_pie_data, cluster_table_data
