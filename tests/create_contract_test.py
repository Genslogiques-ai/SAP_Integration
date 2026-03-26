from sap_core.sap_client import SAPClient
from modules.sd.create_contract import CreateContract

sap = SAPClient()

contract = CreateContract(sap)

contract.execute({
    "customer": "2000",
    "material": "CRUACCE01",
    "quantity": "1000.000"
})