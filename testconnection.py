from sap_core.sap_client import SAPClient

sap = SAPClient()

print(sap.session.Info.SystemName)
print(sap.session.Info.Client)
print(sap.session.Info.User)