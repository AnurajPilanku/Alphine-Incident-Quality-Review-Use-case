import pandas as pd
import sys
x='Agent_PIN,Alert_Status,Area,Assigned_To_Group,Assigned_To_ID,Assigned_To_Name,Assigned_To_PIN,Group Owner PIN,Assigned_To_Site,Category,Case_Type,CI,Closed_By_Id,Closed_By_Name,Closed_By_PIN,Closed_ByGSC,Closed_By_Group,Close_Date,Close Year,Close Month,Close_Time,Closure_Code,Cost_Option,Contact_Id,Contact_Name,Contact_Phone,Contact_PIN,Crisis_Level,Critical_User,Critical_User_Desc,Dept,Description,Column1,Group_Main,Group_Sub,Hand_Offs,Impact,Impact_Desc,Incident_Id,Interaction_Id,Opened_By_Id,Opened_By_Name,Opened_By_PIN,Opened_By_Group,Open_Date,Open Day,Open Month,Open_Days,Open_Time,Open_hour,TimeZone,Outage_End,Outage_Start,Priority,Priority_Desc,Resolve_SLA_Status,Respond_SLA_Status,SLA,SLA_Order,SLA_Breach,SLA_Breach_Resolve,SLA_Breach_Respond,Service,Service_Main,Service_Sub,Service_SLA,Site_Source_Id,Solution,Source,Status,Status_Nbr,Sub_Area,Title,Updated_By_PIN,Update_Time,Update +2,Update_Days,Update_Minutes,Urgency,Urgency_Desc,City,State_Code,Country_Code,Country_Name,Site_Name,IT_Region,IT_Area,ARS_Archive,ARS_Category,ARS_Type,ARS_Item,ERP_Project,Knowledge_Source,Business_Portfolio,Product_Group,Product,Recovery_Tier,ERP_Status,Causing_Service,Vendor_ID,SLA_Used_Time,SLA_Target,Current SLA Status,Current Status,SLA vs Target,Forecasted SLA,Vendor_Reference'
inputexcelname="gsr_data.xlsx"
outputexcelname="aranged_gsr.xlsx"
col_h=x.split(",")
account = pd.read_excel(sys.argv[1] + inputexcelname, engine='openpyxl')
# Arranging columns
account2 = account.reindex(columns=col_h)
#replacing a value in entire excel
account2=account2.replace(['NULL'],'')
account2.to_excel(sys.argv[1] + outputexcelname, index=False)
print("success")
