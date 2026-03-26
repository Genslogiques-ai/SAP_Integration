import time

class CreateContract:

    def __init__(self, sap_client):
        self.sap = sap_client
        self.session = sap_client.session

    # ---------- UTILITIES ----------

    def wait_until_ready(self, timeout=10):
        for _ in range(timeout):
            if not self.session.Busy:
                return
            time.sleep(1)
        raise Exception("SAP not ready")

    def safe_find(self, element_id, retries=3):
        for _ in range(retries):
            try:
                return self.session.findById(element_id)
            except:
                time.sleep(1)
        raise Exception(f"Element not found: {element_id}")

    # ---------- MAIN EXECUTION ----------

    def execute(self, data: dict):
        s = self.session

        # Start VA41
        s.findById("wnd[0]/tbar[0]/okcd").text = "VA41"
        s.findById("wnd[0]").sendVKey(0)
        self.wait_until_ready()

        # Header
        s.findById("wnd[0]/usr/ctxtVBAK-AUART").text = "ZQC"
        s.findById("wnd[0]/usr/ctxtVBAK-VKORG").text = "165s"
        s.findById("wnd[0]/usr/ctxtVBAK-VTWEG").text = "TR"
        s.findById("wnd[0]/usr/ctxtVBAK-SPART").text = "CR"

        s.findById("wnd[0]/tbar[0]/btn[0]").press()
        self.wait_until_ready()

        # Customer
        self.safe_find(
            "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR"
        ).text = "2000"

        self.safe_find(
            "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR"
        ).text = "2000"

        # ✅ CONTROL TAB
        tab = self.safe_find("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01")
        tab.select()
        self.wait_until_ready()

        # Validity dates
        self.safe_find(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4406/"
            "ssubHEADER_FRAME:SAPMV45A:4440/ctxtVBAK-GUEBG"
        ).text = "03/24/2026"

        self.safe_find(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4406/"
            "ssubHEADER_FRAME:SAPMV45A:4440/ctxtVBAK-GUEEN"
        ).text = "03/31/2026"

        # Item table
        table_base = (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4406/"
            "subSUBSCREEN_TC:SAPMV45A:4906/"
            "tblSAPMV45ATCTRL_U_ERF_LIEFERPLAN"
        )

        self.safe_find(f"{table_base}/txtVBAP-POSNR[0,0]").text = "10"
        self.safe_find(f"{table_base}/ctxtRV45A-MABNR[1,0]").text = "CRUACCE01"
        self.safe_find(f"{table_base}/txtVBAP-ZMENG[3,0]").text = "1000.000"

        # ✅ CORRECT INTERACTION (NOT .select())
        mat_field = self.safe_find(f"{table_base}/ctxtRV45A-MABNR[1,0]")
        mat_field.setFocus()
        mat_field.caretPosition = 5

        s.findById("wnd[0]").sendVKey(2)  # ENTER
        self.wait_until_ready()

        # Storage popup handling
        s.findById("wnd[0]").sendVKey(4)
        self.wait_until_ready()

        s.findById("wnd[1]/tbar[0]/btn[0]").press()

        # Incoterms
        s.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\04").select()
        self.wait_until_ready()

        self.safe_find(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\04/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4453/ctxtVBKD-INCO1"
        ).text = "DDP"

        s.findById("wnd[0]/tbar[0]/btn[3]").press()
        s.findById("wnd[1]/tbar[0]/btn[0]").press()

        # TSW Data
        s.findById("wnd[0]").sendVKey(3)

        self.safe_find(
            f"{table_base}/subSUBSCREEN_BUTTONS:SAPMV45A:4052/btnTSW_DATA_PB"
        ).press()

        self.safe_find(
            "wnd[1]/usr/tabsG_TABSTRIP/tabpOIJDET/"
            "ssubSUBSCREEN_BODY:SAPLOIJ_EL_C:0111/"
            "chkOIJ_EL_DOC_MOT-TSWRELV"
        ).selected = True

        self.safe_find(
            "wnd[1]/usr/tabsG_TABSTRIP/tabpOIJDET/"
            "ssubSUBSCREEN_BODY:SAPLOIJ_EL_C:0111/"
            "ctxtOIJ_EL_DOC_MOT-PLANLOC"
        ).text = "GEN_HUS"

        self.safe_find(
            "wnd[1]/usr/tabsG_TABSTRIP/tabpOIJDET/"
            "ssubSUBSCREEN_BODY:SAPLOIJ_EL_C:0111/"
            "ctxtOIJ_EL_DOC_MOT-TSYST"
        ).text = "ATA_MARINE"

        s.findById("wnd[1]/tbar[0]/btn[0]").press()

        # Save
        s.findById("wnd[0]/tbar[0]/btn[11]").press()
        s.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()