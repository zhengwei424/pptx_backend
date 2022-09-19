import pymysql
import json

import urllib3
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
                cluster_namespaces_quota[context][ns]["cpu"] = int(quota.spec.hard.get("limits.cpu"))
                # 单位Gi
                if quota.spec.hard.get("limits.memory")[:-2] == "Gi":
                    cluster_namespaces_quota[context][ns]["memory"] = int(quota.spec.hard.get("limits.memory").split("Gi")[0])
                elif quota.spec.hard.get("limits.memory")[:-2] == "Ti":
                    cluster_namespaces_quota[context][ns]["memory"] = int(quota.spec.hard.get("limits.memory").split("Ti")[0]) * 1024
    # print("deploy_info: ", deploy_info)
    # print("cluster_total_resources: ", cluster_total_resources)
    # print("cluster_namespaces_quota: ", cluster_namespaces_quota)
    # deploy_info = {'fcp': {'apollons-fcp': {'pod-apollo-config-server': 3, 'deployment-apollo-admin-server': 3,
    #                                        'deployment-apollo-portal-server': 1},
    #                       'asmnspace-fcp': {'nginx-web-asm': 2, 'asmview': 2, 'asmservice': 2},
    #                       'bapauthns-fcp': {'sg-collector': 4, 'auth-service': 4, 'unite-codeadmin': 2, 'sg-view': 1,
    #                                         'auth-permission': 2, 'unite-manager': 2, 'sg-gateway': 4,
    #                                         'auth-sec-manager': 2, 'sg-gateway-admin': 4, 'auth-gateway-api': 2,
    #                                         'auth-sec-service': 4, 'sg-cmp-backend': 2},
    #                       'bapnspace-fcp': {'fcpcode': 1, 'pubdata': 3},
    #                       'bpmnspace-fcp': {'nginx-web-bpm': 2, 'bpmservice': 6, 'bpmview': 6},
    #                       'bpsservice-fcp': {'bpsserver': 2, 'bpsworkspace': 1},
    #                       'commander-fcp': {'commander-queue': 2, 'commander-manager': 1, 'commander-pinecone': 2,
    #                                         'commander-bot-manager': 2, 'commander-fe': 2, 'commander-auth': 1,
    #                                         'commander-uic': 2, 'nacos': 1},
    #                       'cpmnspace-fcp': {'nginx-web-cpm': 2, 'cpm-apportion': 2, 'cpm-position': 4, 'deposit': 2,
    #                                         'nginx-web-deposit': 2, 'cpm-autotask': 3, 'cpm-interbank': 4,
    #                                         'cpm-liquidity': 2},
    #                       'crisnspace-fcp': {'crisview': 1, 'nginx-web-cris': 2, 'crisservice': 1},
    #                       'decnspace-fcp': {'nginx-web-dec': 2, 'decisionserver': 1, 'decisionbi': 1}, 'default': {},
    #                       'eastservice-fcp': {'eastservice02': 1, 'eastservice': 1},
    #                       'ecmnspace-fcp': {'ecmscommonservice': 4, 'vimservice': 4, 'nginx-web-vim': 2,
    #                                         'nginx-web-ecms': 2, 'ecmssearchservice': 4, 'ecmsftsservice': 4},
    #                       'egressnginx-fcp': {}, 'esbservice-fcp': {'esb': 3},
    #                       'eurekans-fcp': {'statefulset-eureka-server': 3},
    #                       'fxqservice-fcp': {'nginx-web-aml': 2, 'fxqjobadmin': 1, 'fxqjobexecutor': 1,
    #                                          'fxqservice': 2},
    #                       'gatewayns-fcp': {'sys-file-9698': 4, 'sys-file-9699': 4, 'hellgate-mobile-9692': 4,
    #                                         'hellgate-9697': 4, 'hellgate-9696': 4, 'sys-file-mobile-9693': 4},
    #                       'hexincms-fcp': {'cmsview': 4, 'rdpcw2': 0, 'cmstransservice': 0, 'rdpcw': 0,
    #                                        'nginx-web-cms': 2, 'cmsservice': 3}, 'hexincmv7-fcp': {'v7': 7},
    #                       'hexinecif-fcp': {'nginx-web-ecif': 2, 'ecifview': 5, 'ecifservice': 5},
    #                       'hexinother-fcp': {'asidview': 2, 'ebankview': 4, 'nginx-web-evs': 2, 'asidserivce': 2,
    #                                          'ebankservice': 5, 'nginx-web-ebank': 2, 'evsservice': 4, 'evsview': 4},
    #                       'hexinutip-fcp': {'utipservice': 4, 'utipjobservice': 0, 'utippubservice': 4,
    #                                         'bedc-service': 3, 'utipfileservice': 4},
    #                       'ingress-nginx': {'default-http-backend': 0}, 'kube-node-lease': {}, 'kube-public': {},
    #                       'kube-system': {}, 'kubernetes-dashboard': {}, 'mapns': {},
    #                       'msgcenter-fcp': {'msgcenterservice': 2, 'msgcenterserviceconsumer': 2},
    #                       'nfbcmnspace-fcp': {'nfbcm-message-service': 2, 'nfbcm-message-front': 2,
    #                                           'nfbcm-service-fbcm': 2, 'nfbcm-greport-service': 2,
    #                                           'nfbcm-ebs-service': 2, 'nfbcm-service-crm': 2, 'nfbcm-service-ics': 2,
    #                                           'nginx-web-nfbcm': 2, 'nfbcm-service-prms': 2,
    #                                           'nfbcm-app-center-service': 2, 'nfbcm-data-transfer': 2,
    #                                           'nfbcm-service-imes': 2, 'nfbcm-neams': 2, 'nfbcm-service-adapter': 2},
    #                       'nginxgate-fcp': {'nginx-web-mobile': 2, 'nginx-web-gate': 2},
    #                       'odscnspace-fcp': {'odsctlmagent': 1, 'odsctlmain': 1, 'odsctlsagent2': 1,
    #                                          'odsctlsagent1': 1},
    #                       'pubnspace-fcp': {'pubview': 6, 'nginx-web-organizat': 2, 'pubservice': 6},
    #                       'treasurer-fcp': {'treasurer-acct-service': 4, 'fcp-archive': 1, 'nginx-web-siku': 2,
    #                                         'treasurer-file': 4, 'treasurer-pay-view': 6, 'treasurer-pay-service': 6,
    #                                         'treasurer-fin-view': 2, 'treasurer-pos-view': 4, 'nginx-web-dv': 2,
    #                                         'treasurer-pos-job': 1, 'treasurer-dv-view': 2, 'treasurer-pay-job': 1,
    #                                         'treasurer-dv-job': 1, 'treasurer-bud-view': 6, 'treasurer-acct-view': 4,
    #                                         'treasurer-dv-service': 2, 'treasurer-fin-service': 2,
    #                                         'treasurer-fin-job': 1, 'treasurer-bud-service': 6, 'treasurer-acct-job': 1,
    #                                         'nginx-web-finance': 2, 'treasurer-bud-job': 1, 'treasurer-pos-service': 4},
    #                       'uapservice-fcp': {'uapservice': 3},
    #                       'xquantns-fcp': {'xquant-calc': 0, 'xquant-smartbi': 0, 'xquant-xir': 0},
    #                       'xxljobns-fcp': {'xxljob': 3}, 'zabbixns-fcp': {'airflow': 1, 'monitor': 1}},
    #               'fcp-inner-microservice': {
    #                   '01fd132b9a74487cb3a423d5aa74ff2e-fcp-inner-microservice': {'cpm-apportion': 2,
    #                                                                               'cpm-liquidity': 2,
    #                                                                               'cpm-interbank': 4, 'cpm-autotask': 3,
    #                                                                               'cpm-position': 4},
    #                   '04fadaa5cb654b97a45d6c2b79242391-fcp-inner-microservice': {'uapservice': 4},
    #                   '65e4622647f74ed2b815265fb6fc0f96-fcp-inner-microservice': {'nfbcm-service-adapter': 2,
    #                                                                               'nfbcm-data-transfer': 2,
    #                                                                               'nfbcm-service-imes': 2,
    #                                                                               'nfbcm-neams': 2,
    #                                                                               'nfbcm-service-prms': 2,
    #                                                                               'nfbcm-message-front': 2,
    #                                                                               'nfbcm-service-fbcm': 2,
    #                                                                               'nfbcm-greport-service': 2,
    #                                                                               'nfbcm-ebs-service': 2,
    #                                                                               'nfbcm-app-center-service': 2,
    #                                                                               'nfbcm-message-service': 2,
    #                                                                               'nfbcm-service-ics': 2,
    #                                                                               'nfbcm-service-crm': 2},
    #                   '81f4126e11b74bf3a5c7cd64dfb032d0-fcp-inner-microservice': {'asidserivce': 2, 'evsservice': 5,
    #                                                                               'ebankservice': 5,
    #                                                                               'cmstransservice': 1, 'cmsservice': 4,
    #                                                                               'ebankview': 4, 'asidview': 2,
    #                                                                               'cmsview': 4, 'utipfileservice': 4,
    #                                                                               'utippubservice': 4, 'utipservice': 4,
    #                                                                               'utipjobservice': 1,
    #                                                                               'bedcservice': 4},
    #                   'ae3ce9d37f3241d79406ba7f3fe979f7-fcp-inner-microservice': {'ecifservice': 5}, 'default': {},
    #                   'f7e35b2438ec4029b713f0aa5e4a19c7-fcp-inner-microservice': {'msgcenterservice': 2,
    #                                                                               'msgcenterserviceconsumer': 2},
    #                   'ingress-nginx': {'default-http-backend': 1}, 'kube-public': {}, 'kube-system': {},
    #                   'monitoring': {'grafana': 1, 'kube-state-metrics': 2, 'prometheus': 1}}, 'fcp-inner-backend': {
    #         '3c2558da75bf4e0580ec305b8b80c510-fcp-inner-backend': {'odsctlsagent1': 0, 'odsctlmain': 0,
    #                                                                'odsctlsagent2': 0, 'odsctlmagent': 0},
    #         '81f4126e11b74bf3a5c7cd64dfb032d0-fcp-inner-backend': {'rdpcw': 1, 'rdpcw2': 1, 'v7cw': 7},
    #         'cbf3dc9611594e7e962dcbfa928448a9-fcp-inner-backend': {'commander-pinecone': 1, 'nacos': 0,
    #                                                                'commander-bot-manager': 2, 'commander-manager': 1,
    #                                                                'commander-queue': 2, 'commander-uic': 2,
    #                                                                'commander-auth': 1, 'commander-fe': 2},
    #         'default': {}, 'ingress-nginx': {'default-http-backend': 0}, 'kube-public': {}, 'kube-system': {},
    #         'monitoring': {'grafana': 1, 'kube-state-metrics': 2, 'prometheus': 1}}, 'fcp-outer-microservice': {
    #         '10b5d426dea1448b8ac3fad727eda2dc-fcp-outer-microservice': {'bpmservice': 6},
    #         '66df708708474b34a459bf04c8a73cc2-fcp-outer-microservice': {'secretkey-service': 4, 'fcpcode-admin': 2,
    #                                                                     'uims-gateway-api': 2, 'deposit': 2,
    #                                                                     'sg-gateway-admin': 4, 'uims-service': 4,
    #                                                                     'sg-gateway': 4, 'sg-collector': 4,
    #                                                                     'uims-permission': 2, 'secretkey-manager': 2,
    #                                                                     'sg-view': 1, 'sg-cmp-backend': 2, 'fcpcode': 5,
    #                                                                     'pubdata': 3, 'conf-manager': 2},
    #         '6c142e16264740489eb87750d6aa14a4-fcp-outer-microservice': {'crisservice': 0, 'crisview': 0},
    #         '6cd297e121f148fab73f0090417edbe3-fcp-outer-microservice': {'asmview': 2, 'asmservice': 2},
    #         '6f1c55ccff034b099c5323b9f1df21c5-fcp-outer-microservice': {'bpmview': 6},
    #         '7505e75cfbc74e3b81b2592273a540f3-fcp-outer-microservice': {'pubservice': 6},
    #         '82a1880da1b544a582cc6e133febc3b8-fcp-outer-microservice': {'treasurer-acct-view': 4,
    #                                                                     'treasurer-dv-view': 2,
    #                                                                     'treasurer-pay-service': 6,
    #                                                                     'treasurer-pos-service': 4,
    #                                                                     'treasurer-pos-view': 4,
    #                                                                     'treasurer-bud-service': 6,
    #                                                                     'treasurer-fin-view': 2,
    #                                                                     'treasurer-fin-service': 2,
    #                                                                     'treasurer-pay-view': 6,
    #                                                                     'treasurer-dv-service': 2,
    #                                                                     'treasurer-bud-view': 6,
    #                                                                     'treasurer-acct-service': 4,
    #                                                                     'treasurer-file': 2},
    #         '9272869c52a14468a7aa62b5dd9b39b2-fcp-outer-microservice': {'evsview': 4},
    #         '9efff7f47f7b4f7583127851146cac7d-fcp-outer-microservice': {'pubview': 6},
    #         'af7a468b5ae94d1d9a09b21d07f6876c-fcp-outer-microservice': {'ecifview': 5},
    #         'c95216c44f6e406997390bd908fff962-fcp-outer-microservice': {'ecmssearchservice': 4, 'vimservice': 4,
    #                                                                     'ecmsftsservice': 4, 'ecmscommonservice': 4},
    #         'default': {}, 'ingress-nginx': {'default-http-backend': 1}, 'kube-public': {}, 'kube-system': {},
    #         'monitoring': {'grafana': 1, 'kube-state-metrics': 2, 'prometheus': 1}}, 'fcp-outer-backend': {
    #         '149f389b3fec4d17a80d077f84c486dd-fcp-outer-backend': {'esb-server3': 1, 'esb-server2': 1, 'esb-server1': 1,
    #                                                                'esb-dangban': 0},
    #         '82a1880da1b544a582cc6e133febc3b8-fcp-outer-backend': {'fcp-archive': 1, 'treasurer-pay-job': 1,
    #                                                                'treasurer-pos-job': 1, 'treasurer-acct-job': 1,
    #                                                                'treasurer-fin-job': 1, 'treasurer-bud-job': 1,
    #                                                                'treasurer-dv-job': 1},
    #         '9a64cea0e26d4382a8f9395a71fc0b43-fcp-outer-backend': {'eastservice02': 0, 'eastservice': 0},
    #         'a542d960d3bc4f4aa1a5b193291c2998-fcp-outer-backend': {'decisionbi': 0, 'decisionserver': 0},
    #         'aee3b90ba924463bab7d83c6197f807e-fcp-outer-backend': {'fxqjobadmin': 0, 'fxqjobexecutor': 0,
    #                                                                'fxqservice': 1, 'fxqservice2': 1},
    #         'b8d62b960c51456898841d831cf65df0-fcp-outer-backend': {},
    #         'c82194ba3af84e06a0b7a6413fb31cf2-fcp-outer-backend': {'bpsworkspace': 1, 'bpsserver2': 1, 'bpsserver': 1},
    #         'dcb8a27ee8064ee78f3fcbdbbabcce0f-fcp-outer-backend': {'xquant-xir': 1, 'xquant-xir2': 1, 'xquant-calc2': 1,
    #                                                                'xquant-calc': 1, 'xquant-smartbi': 1},
    #         'default': {}, 'ingress-nginx': {'default-http-backend': 1}, 'kube-public': {}, 'kube-system': {},
    #         'monitoring': {'grafana': 1, 'kube-state-metrics': 2, 'prometheus': 1}}
    #      }
    # cluster_total_resources= {'fcp': {'cpu': 1856, 'memory': 12057},
    #                           'fcp-inner-microservice': {'cpu': 1024, 'memory': 4028},
    #                           'fcp-inner-backend': {'cpu': 680, 'memory': 2768},
    #                           'fcp-outer-microservice': {'cpu': 1536, 'memory': 6041},
    #                           'fcp-outer-backend': {'cpu': 384, 'memory': 1495}}
    # cluster_namespaces_quota = {
    #     'fcp': {'apollons-fcp': {'cpu': 15, 'memory': 0}, 'asmnspace-fcp': {'cpu': 11, 'memory': 0},
    #             'bapauthns-fcp': {'cpu': 66, 'memory': 0}, 'bapnspace-fcp': {'cpu': 17, 'memory': 0},
    #             'bpmnspace-fcp': {'cpu': 51, 'memory': 0}, 'bpsservice-fcp': {'cpu': 13, 'memory': 0},
    #             'commander-fcp': {'cpu': 29, 'memory': 0}, 'cpmnspace-fcp': {'cpu': 44, 'memory': 0},
    #             'crisnspace-fcp': {'cpu': 8, 'memory': 0}, 'decnspace-fcp': {'cpu': 19, 'memory': 0},
    #             'default': {'cpu': 0, 'memory': 0}, 'eastservice-fcp': {'cpu': 9, 'memory': 0},
    #             'ecmnspace-fcp': {'cpu': 101, 'memory': 0}, 'egressnginx-fcp': {'cpu': 7, 'memory': 0},
    #             'esbservice-fcp': {'cpu': 25, 'memory': 0}, 'eurekans-fcp': {'cpu': 9, 'memory': 0},
    #             'fxqservice-fcp': {'cpu': 19, 'memory': 0}, 'gatewayns-fcp': {'cpu': 193, 'memory': 0},
    #             'hexincms-fcp': {'cpu': 91, 'memory': 0}, 'hexincmv7-fcp': {'cpu': 57, 'memory': 0},
    #             'hexinecif-fcp': {'cpu': 23, 'memory': 0}, 'hexinother-fcp': {'cpu': 85, 'memory': 0},
    #             'hexinutip-fcp': {'cpu': 133, 'memory': 0}, 'ingress-nginx': {'cpu': 0, 'memory': 0},
    #             'kube-node-lease': {'cpu': 0, 'memory': 0}, 'kube-public': {'cpu': 0, 'memory': 0},
    #             'kube-system': {'cpu': 0, 'memory': 0}, 'kubernetes-dashboard': {'cpu': 0, 'memory': 0},
    #             'mapns': {'cpu': 0, 'memory': 0}, 'msgcenter-fcp': {'cpu': 17, 'memory': 0},
    #             'nfbcmnspace-fcp': {'cpu': 123, 'memory': 0}, 'nginxgate-fcp': {'cpu': 9, 'memory': 0},
    #             'odscnspace-fcp': {'cpu': 17, 'memory': 0}, 'pubnspace-fcp': {'cpu': 51, 'memory': 0},
    #             'treasurer-fcp': {'cpu': 363, 'memory': 0}, 'uapservice-fcp': {'cpu': 29, 'memory': 0},
    #             'xquantns-fcp': {'cpu': 73, 'memory': 0}, 'xxljobns-fcp': {'cpu': 7, 'memory': 0},
    #             'zabbixns-fcp': {'cpu': 17, 'memory': 0}},
    #     'fcp-inner-microservice': {'01fd132b9a74487cb3a423d5aa74ff2e-fcp-inner-microservice': {'cpu': 40, 'memory': 0},
    #                                '04fadaa5cb654b97a45d6c2b79242391-fcp-inner-microservice': {'cpu': 35, 'memory': 0},
    #                                '65e4622647f74ed2b815265fb6fc0f96-fcp-inner-microservice': {'cpu': 128, 'memory': 0},
    #                                '81f4126e11b74bf3a5c7cd64dfb032d0-fcp-inner-microservice': {'cpu': 284, 'memory': 0},
    #                                'ae3ce9d37f3241d79406ba7f3fe979f7-fcp-inner-microservice': {'cpu': 16, 'memory': 0},
    #                                'default': {'cpu': 0, 'memory': 0},
    #                                'f7e35b2438ec4029b713f0aa5e4a19c7-fcp-inner-microservice': {'cpu': 24, 'memory': 0},
    #                                'ingress-nginx': {'cpu': 0, 'memory': 0}, 'kube-public': {'cpu': 0, 'memory': 0},
    #                                'kube-system': {'cpu': 0, 'memory': 0}, 'monitoring': {'cpu': 0, 'memory': 0}},
    #     'fcp-inner-backend': {'3c2558da75bf4e0580ec305b8b80c510-fcp-inner-backend': {'cpu': 16, 'memory': 0},
    #                           '81f4126e11b74bf3a5c7cd64dfb032d0-fcp-inner-backend': {'cpu': 112, 'memory': 0},
    #                           'cbf3dc9611594e7e962dcbfa928448a9-fcp-inner-backend': {'cpu': 35, 'memory': 0},
    #                           'default': {'cpu': 0, 'memory': 0}, 'ingress-nginx': {'cpu': 0, 'memory': 0},
    #                           'kube-public': {'cpu': 0, 'memory': 0}, 'kube-system': {'cpu': 0, 'memory': 0},
    #                           'monitoring': {'cpu': 0, 'memory': 0}},
    #     'fcp-outer-microservice': {'10b5d426dea1448b8ac3fad727eda2dc-fcp-outer-microservice': {'cpu': 28, 'memory': 0},
    #                                '66df708708474b34a459bf04c8a73cc2-fcp-outer-microservice': {'cpu': 95, 'memory': 0},
    #                                '6c142e16264740489eb87750d6aa14a4-fcp-outer-microservice': {'cpu': 10, 'memory': 0},
    #                                '6cd297e121f148fab73f0090417edbe3-fcp-outer-microservice': {'cpu': 12, 'memory': 0},
    #                                '6f1c55ccff034b099c5323b9f1df21c5-fcp-outer-microservice': {'cpu': 28, 'memory': 0},
    #                                '7505e75cfbc74e3b81b2592273a540f3-fcp-outer-microservice': {'cpu': 28, 'memory': 0},
    #                                '82a1880da1b544a582cc6e133febc3b8-fcp-outer-microservice': {'cpu': 513, 'memory': 0},
    #                                '9272869c52a14468a7aa62b5dd9b39b2-fcp-outer-microservice': {'cpu': 20, 'memory': 0},
    #                                '9efff7f47f7b4f7583127851146cac7d-fcp-outer-microservice': {'cpu': 28, 'memory': 0},
    #                                'af7a468b5ae94d1d9a09b21d07f6876c-fcp-outer-microservice': {'cpu': 20, 'memory': 0},
    #                                'c95216c44f6e406997390bd908fff962-fcp-outer-microservice': {'cpu': 120, 'memory': 0},
    #                                'default': {'cpu': 0, 'memory': 0}, 'ingress-nginx': {'cpu': 0, 'memory': 0},
    #                                'kube-public': {'cpu': 0, 'memory': 0}, 'kube-system': {'cpu': 0, 'memory': 0},
    #                                'monitoring': {'cpu': 0, 'memory': 0}},
    #     'fcp-outer-backend': {'149f389b3fec4d17a80d077f84c486dd-fcp-outer-backend': {'cpu': 24, 'memory': 0},
    #                           '82a1880da1b544a582cc6e133febc3b8-fcp-outer-backend': {'cpu': 72, 'memory': 0},
    #                           '9a64cea0e26d4382a8f9395a71fc0b43-fcp-outer-backend': {'cpu': 10, 'memory': 0},
    #                           'a542d960d3bc4f4aa1a5b193291c2998-fcp-outer-backend': {'cpu': 16, 'memory': 0},
    #                           'aee3b90ba924463bab7d83c6197f807e-fcp-outer-backend': {'cpu': 17, 'memory': 0},
    #                           'b8d62b960c51456898841d831cf65df0-fcp-outer-backend': {'cpu': 90, 'memory': 0},
    #                           'c82194ba3af84e06a0b7a6413fb31cf2-fcp-outer-backend': {'cpu': 24, 'memory': 0},
    #                           'dcb8a27ee8064ee78f3fcbdbbabcce0f-fcp-outer-backend': {'cpu': 72, 'memory': 0},
    #                           'default': {'cpu': 0, 'memory': 0}, 'ingress-nginx': {'cpu': 0, 'memory': 0},
    #                           'kube-public': {'cpu': 0, 'memory': 0}, 'kube-system': {'cpu': 0, 'memory': 0},
    #                           'monitoring': {'cpu': 0, 'memory': 0}}}

    return deploy_info, cluster_total_resources, cluster_namespaces_quota


def get_namespace_project():
    with open('config/config.json', encoding='utf-8') as config:
        config = json.load(config)
    conn = pymysql.connect(**(config["mysql"]))
    mh = MySQLHandler(conn)
    sql = "select PROJECT_ID, PROJECT_NAME from dps_pjm_project"
    result = mh.select(sql)
    return result


if __name__ == '__main__':
    result = (
        ('01fd132b9a74487cb3a423d5aa74ff2e', '财务公司头寸管理'), ('04fadaa5cb654b97a45d6c2b79242391', '统一接入平台'),
        ('10b5d426dea1448b8ac3fad727eda2dc', 'bpm-审批域中台'), ('149f389b3fec4d17a80d077f84c486dd', 'ESB'),
        ('1cb357a1fe9642268d4aa09cdfb91a3c', 'bpm'), ('37badd726f694af4868785bd4d964878', '金融服务平台前端'),
        ('3c2558da75bf4e0580ec305b8b80c510', '数据仓库'), ('63bf2873c8d948038ef49d045c252fcb', '统一接入'),
        ('65e4622647f74ed2b815265fb6fc0f96', '信贷'), ('66df708708474b34a459bf04c8a73cc2', '平台公共服务'),
        ('6c142e16264740489eb87750d6aa14a4', '监管报送'), ('6cd297e121f148fab73f0090417edbe3', 'asm事后监督系统'),
        ('6f1c55ccff034b099c5323b9f1df21c5', 'bpm-审批域前台'), ('7505e75cfbc74e3b81b2592273a540f3', '公共服务'),
        ('81f4126e11b74bf3a5c7cd64dfb032d0', '核心'), ('82a1880da1b544a582cc6e133febc3b8', '司库'),
        ('9272869c52a14468a7aa62b5dd9b39b2', '电子凭证前台'), ('9a64cea0e26d4382a8f9395a71fc0b43', 'east数据报送'),
        ('9efff7f47f7b4f7583127851146cac7d', '公共管理'), ('a542d960d3bc4f4aa1a5b193291c2998', '决策分析'),
        ('ae3ce9d37f3241d79406ba7f3fe979f7', '客户信息管理系统'), ('aee3b90ba924463bab7d83c6197f807e', '反洗钱'),
        ('af7a468b5ae94d1d9a09b21d07f6876c', '客户信息管理前台'), ('apollons', '配置中心-绿区'),
        ('asmnspace', 'asm事后监督系统-绿区'), ('b8d62b960c51456898841d831cf65df0', '智慧客服系统'),
        ('bapauthns', '平台服务-绿区'), ('bapnspace', '平台基础服务-绿区'), ('bpmnspace', 'bpm流程引擎-绿区'),
        ('bpsservice', 'BPS流程引擎-绿区'), ('c82194ba3af84e06a0b7a6413fb31cf2', 'bps-流程引擎后台'),
        ('c95216c44f6e406997390bd908fff962', '内容影像'), ('cbf3dc9611594e7e962dcbfa928448a9', 'RPA'),
        ('commander', 'rpa机器人-绿区'), ('cpmnspace', 'cpm财务公司头寸-绿区'), ('crisnspace', '监管报送-绿区'),
        ('dcb8a27ee8064ee78f3fcbdbbabcce0f', '投资管理系统'), ('decnspace', '决策分析-绿区'),
        ('eastservice', 'east数据报送-绿区'), ('ecmnspace', '内容影像-绿区'), ('egressnginx', '出口网关服务-绿区'),
        ('esbservice', 'esb服务总线-绿区'), ('eurekans', '注册中心-绿区'),
        ('f7e35b2438ec4029b713f0aa5e4a19c7', '消息中心'),
        ('fxqservice', '反洗钱-绿区'), ('gatewayns', '应用网关-绿区'), ('hexincms', '核心-结算-绿区'),
        ('hexincmv7', '核心v7服务-绿区'), ('hexinecif', '核心-客户信息管理-绿区'), ('hexinother', '核心-其他-绿区'),
        ('hexinutip', '核心-银企-绿区'), ('msgcenter', '消息中心-绿区'), ('nfbcmnspace', 'nfbcm新信贷-绿区'),
        ('nginxgate', '前端框架路由分发-绿区'), ('odscnspace', 'ods数据仓库-绿区'), ('pubnspace', '公共管理-绿区'),
        ('treasurer', '集团司库-绿区'), ('uapservice', '统一接入平台-绿区'), ('xquantns', '投资-绿区'),
        ('xxljobns', 'xxljob调度管理-绿区'), ('zabbixns', '运维巡检监控-绿区'))
    prj = {}
    for item in result:
        prj[item[0]] = item[1]
    get_deploy_info()
    data = {'fcp': {'apollons-fcp': {'': 1}, 'asmnspace-fcp': {'asmview': 2, 'asmservice': 2},
                    'bapauthns-fcp': {'sg-collector': 4, 'auth-service': 4, 'unite-codeadmin': 2, 'sg-view': 1,
                                      'auth-permission': 2, 'unite-manager': 2, 'sg-gateway': 4, 'auth-sec-manager': 2,
                                      'sg-gateway-admin': 4, 'auth-gateway-api': 2, 'auth-sec-service': 4,
                                      'sg-cmp-backend': 2}, 'bapnspace-fcp': {'fcpcode': 1, 'pubdata': 3},
                    'bpmnspace-fcp': {'bpmservice': 6, 'bpmview': 6}, 'bpsservice-fcp': {'': 1},
                    'commander-fcp': {'commander-queue': 2, 'commander-manager': 1, 'commander-pinecone': 2,
                                      'commander-bot-manager': 2, 'commander-fe': 2, 'commander-auth': 1,
                                      'commander-uic': 2, 'nacos': 1},
                    'cpmnspace-fcp': {'cpm-apportion': 2, 'cpm-position': 4, 'deposit': 2, 'cpm-autotask': 3,
                                      'cpm-interbank': 4, 'cpm-liquidity': 2},
                    'crisnspace-fcp': {'crisview': 1, 'crisservice': 1},
                    'decnspace-fcp': {'decisionserver': 1, 'decisionbi': 1}, 'default': {},
                    'eastservice-fcp': {'eastservice02': 1, 'eastservice': 1},
                    'ecmnspace-fcp': {'ecmscommonservice': 4, 'vimservice': 4, 'ecmssearchservice': 4,
                                      'ecmsftsservice': 4}, 'egressnginx-fcp': {}, 'esbservice-fcp': {},
                    'eurekans-fcp': {}, 'fxqservice-fcp': {'': 2},
                    'gatewayns-fcp': {'sys-file-9698': 4, 'sys-file-9699': 4, 'hellgate-mobile-9692': 4,
                                      'hellgate-9697': 4, 'hellgate-9696': 4, 'sys-file-mobile-9693': 4},
                    'hexincms-fcp': {'cmsview': 4, 'rdpcw2': 0, 'cmstransservice': 0, 'rdpcw': 0, '': 6},
                    'hexincmv7-fcp': {}, 'hexinecif-fcp': {'ecifview': 5, 'ecifservice': 5},
                    'hexinother-fcp': {'asidview': 2, 'ebankview': 4, 'asidserivce': 2, 'ebankservice': 5, '': 4,
                                       'evsview': 4},
                    'hexinutip-fcp': {'utipservice': 4, 'utipjobservice': 1, 'utippubservice': 4, '': 4},
                    'ingress-nginx': {'': 0}, 'kube-node-lease': {}, 'kube-public': {}, 'kube-system': {'': 1},
                    'kubernetes-dashboard': {'': 1}, 'mapns': {}, 'msgcenter-fcp': {'': 2},
                    'nfbcmnspace-fcp': {'nfbcm-message-service': 2, 'nfbcm-message-front': 2, 'nfbcm-service-fbcm': 2,
                                        'nfbcm-greport-service': 2, 'nfbcm-ebs-service': 2, 'nfbcm-service-crm': 2,
                                        'nfbcm-service-ics': 2, 'nfbcm-service-prms': 2, 'nfbcm-app-center-service': 2,
                                        'nfbcm-data-transfer': 2, 'nfbcm-service-imes': 2, 'nfbcm-neams': 2,
                                        'nfbcm-service-adapter': 2}, 'nginxgate-fcp': {'nginx-web-gate': 2},
                    'odscnspace-fcp': {'odsctlmagent': 1, 'odsctlmain': 1, 'odsctlsagent2': 1, 'odsctlsagent1': 1},
                    'pubnspace-fcp': {'pubview': 6, 'pubservice': 6},
                    'treasurer-fcp': {'treasurer-acct-service': 4, 'fcp-archive': 1, 'treasurer-file': 4,
                                      'treasurer-pay-view': 6, 'treasurer-pay-service': 6, 'treasurer-fin-view': 2,
                                      'treasurer-pos-view': 4, 'treasurer-pos-job': 1, 'treasurer-dv-view': 2,
                                      'treasurer-pay-job': 1, 'treasurer-dv-job': 1, 'treasurer-bud-view': 6,
                                      'treasurer-acct-view': 4, 'treasurer-dv-service': 2, 'treasurer-fin-service': 2,
                                      'treasurer-fin-job': 1, 'treasurer-bud-service': 6, 'treasurer-acct-job': 1,
                                      'treasurer-bud-job': 1, 'treasurer-pos-service': 4}, 'uapservice-fcp': {'': 6},
                    'xquantns-fcp': {}, 'xxljobns-fcp': {'xxljob': 3}, 'zabbixns-fcp': {'': 1}},
            'fcp-inner-microservice': {
                '01fd132b9a74487cb3a423d5aa74ff2e-fcp-inner-microservice': {'cpm-apportion': 2, 'cpm-liquidity': 2,
                                                                            'cpm-interbank': 4, 'cpm-autotask': 3,
                                                                            'cpm-position': 4},
                '04fadaa5cb654b97a45d6c2b79242391-fcp-inner-microservice': {'uapservice': 6},
                '65e4622647f74ed2b815265fb6fc0f96-fcp-inner-microservice': {'nfbcm-service-adapter': 2,
                                                                            'nfbcm-data-transfer': 2,
                                                                            'nfbcm-service-imes': 2, 'nfbcm-neams': 2,
                                                                            'nfbcm-service-prms': 2,
                                                                            'nfbcm-message-front': 2,
                                                                            'nfbcm-service-fbcm': 2,
                                                                            'nfbcm-greport-service': 2,
                                                                            'nfbcm-ebs-service': 2,
                                                                            'nfbcm-app-center-service': 2,
                                                                            'nfbcm-message-service': 2,
                                                                            'nfbcm-service-ics': 2,
                                                                            'nfbcm-service-crm': 2},
                '81f4126e11b74bf3a5c7cd64dfb032d0-fcp-inner-microservice': {'asidserivce': 2, 'evsservice': 5,
                                                                            'ebankservice': 5, 'cmstransservice': 1,
                                                                            'cmsservice': 6, 'ebankview': 4,
                                                                            'asidview': 2, 'cmsview': 4,
                                                                            'utipfileservice': 4, 'utippubservice': 4,
                                                                            'utipservice': 4, 'utipjobservice': 1,
                                                                            'bedcservice': 4},
                'ae3ce9d37f3241d79406ba7f3fe979f7-fcp-inner-microservice': {'ecifservice': 5}, 'default': {},
                'f7e35b2438ec4029b713f0aa5e4a19c7-fcp-inner-microservice': {'msgcenterservice': 2,
                                                                            'msgcenterserviceconsumer': 2},
                'ingress-nginx': {'': 1}, 'kube-public': {}, 'kube-system': {'': 1}, 'monitoring': {'': 1}},
            'fcp-inner-backend': {
                '3c2558da75bf4e0580ec305b8b80c510-fcp-inner-backend': {'odsctlsagent1': 0, 'odsctlmain': 0,
                                                                       'odsctlsagent2': 0, 'odsctlmagent': 0},
                '81f4126e11b74bf3a5c7cd64dfb032d0-fcp-inner-backend': {'rdpcw': 1, 'rdpcw2': 1, 'v7cw': 7},
                'cbf3dc9611594e7e962dcbfa928448a9-fcp-inner-backend': {'commander-pinecone': 1, 'nacos': 0,
                                                                       'commander-bot-manager': 2,
                                                                       'commander-manager': 1, 'commander-queue': 2,
                                                                       'commander-uic': 2, 'commander-auth': 1,
                                                                       'commander-fe': 2}, 'default': {},
                'ingress-nginx': {'': 0}, 'kube-public': {}, 'kube-system': {'': 1}, 'monitoring': {'': 1}},
            'fcp-outer-microservice': {'10b5d426dea1448b8ac3fad727eda2dc-fcp-outer-microservice': {'bpmservice': 6},
                                       '66df708708474b34a459bf04c8a73cc2-fcp-outer-microservice': {
                                           'secretkey-service': 4, 'fcpcode-admin': 2, 'uims-gateway-api': 2,
                                           'deposit': 2, 'sg-gateway-admin': 4, 'uims-service': 4, 'sg-gateway': 4,
                                           'sg-collector': 4, 'uims-permission': 2, 'secretkey-manager': 2,
                                           'sg-view': 1, 'sg-cmp-backend': 2, 'fcpcode': 5, 'pubdata': 3,
                                           'conf-manager': 2},
                                       '6c142e16264740489eb87750d6aa14a4-fcp-outer-microservice': {'crisservice': 0,
                                                                                                   'crisview': 0},
                                       '6cd297e121f148fab73f0090417edbe3-fcp-outer-microservice': {'asmview': 2,
                                                                                                   'asmservice': 2},
                                       '6f1c55ccff034b099c5323b9f1df21c5-fcp-outer-microservice': {'bpmview': 6},
                                       '7505e75cfbc74e3b81b2592273a540f3-fcp-outer-microservice': {'pubservice': 6},
                                       '82a1880da1b544a582cc6e133febc3b8-fcp-outer-microservice': {
                                           'treasurer-acct-view': 4, 'treasurer-dv-view': 2, 'treasurer-pay-service': 6,
                                           'treasurer-pos-service': 4, 'treasurer-pos-view': 4,
                                           'treasurer-bud-service': 6, 'treasurer-fin-view': 2,
                                           'treasurer-fin-service': 2, 'treasurer-pay-view': 6,
                                           'treasurer-dv-service': 2, 'treasurer-bud-view': 6,
                                           'treasurer-acct-service': 4, 'treasurer-file': 2},
                                       '9272869c52a14468a7aa62b5dd9b39b2-fcp-outer-microservice': {'evsview': 4},
                                       '9efff7f47f7b4f7583127851146cac7d-fcp-outer-microservice': {'pubview': 6},
                                       'af7a468b5ae94d1d9a09b21d07f6876c-fcp-outer-microservice': {'ecifview': 5},
                                       'c95216c44f6e406997390bd908fff962-fcp-outer-microservice': {
                                           'ecmssearchservice': 4, 'vimservice': 4, 'ecmsftsservice': 4,
                                           'ecmscommonservice': 4}, 'default': {}, 'ingress-nginx': {'': 1},
                                       'kube-public': {}, 'kube-system': {'': 1}, 'monitoring': {'': 1}},
            'fcp-outer-backend': {
                '149f389b3fec4d17a80d077f84c486dd-fcp-outer-backend': {'esb-server3': 1, 'esb-server2': 1,
                                                                       'esb-server1': 1, 'esb-dangban': 0},
                '82a1880da1b544a582cc6e133febc3b8-fcp-outer-backend': {'fcp-archive': 1, 'treasurer-pay-job': 1,
                                                                       'treasurer-pos-job': 1, 'treasurer-acct-job': 1,
                                                                       'treasurer-fin-job': 1, 'treasurer-bud-job': 1,
                                                                       'treasurer-dv-job': 1},
                '9a64cea0e26d4382a8f9395a71fc0b43-fcp-outer-backend': {'eastservice02': 0, 'eastservice': 0},
                'a542d960d3bc4f4aa1a5b193291c2998-fcp-outer-backend': {'decisionbi': 0, 'decisionserver': 0},
                'aee3b90ba924463bab7d83c6197f807e-fcp-outer-backend': {'fxqjobadmin': 0, 'fxqjobexecutor': 0,
                                                                       'fxqservice': 1, 'fxqservice2': 1},
                'b8d62b960c51456898841d831cf65df0-fcp-outer-backend': {},
                'c82194ba3af84e06a0b7a6413fb31cf2-fcp-outer-backend': {'bpsworkspace': 1, 'bpsserver2': 1,
                                                                       'bpsserver': 1},
                'dcb8a27ee8064ee78f3fcbdbbabcce0f-fcp-outer-backend': {'xquant-xir': 1, 'xquant-xir2': 1,
                                                                       'xquant-calc2': 1, 'xquant-calc': 1,
                                                                       'xquant-smartbi': 1}, 'default': {},
                'ingress-nginx': {'': 1}, 'kube-public': {}, 'kube-system': {'': 1}, 'monitoring': {'': 1}}
            }
    # print(json.dumps(deploy_info, indent=2))
    # deploy_info = {}
    # for key in data.keys():
    #     deploy_info[key] = [["项目名称", "服务名称", "实例数"]]
    #     for ns in data.get(key).keys():
    #         postfix = ns.split('-')[-1]
    #         if postfix == 'fcp' or postfix == 'microservice' or postfix == 'backend':
    #             for replica_info in data.get(key).get(ns).items():
    #                 tmp = []
    #                 for prj_id in prj.keys():
    #                     if prj_id in ns:
    #                         tmp.append(prj[prj_id])
    #                 tmp += list(replica_info)
    #                 deploy_info[key].append(tmp)
    #     # 合计行
    #     total = []
    #     # 统计项目数
    #     ns_temp = []
    #     # 统计应用副本数
    #     replicas_count = 0
    #     for index in range(1, len(deploy_info[key])):
    #         replicas_count += int(deploy_info[key][index][2])
    #         ns_temp.append(deploy_info[key][index][0])
    #     ns_temp = list(set(ns_temp))
    #     total.append(str(len(ns_temp)))
    #     # 统计应用数
    #     total.append(str(len(deploy_info[key]) - 1))
    #     total.append(str(replicas_count))
    #     deploy_info[key].append(total)
    # print(json.dumps(deploy_info))
