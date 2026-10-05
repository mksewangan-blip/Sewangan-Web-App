
function repairSewanganAndroidV2(){
 const ss=SpreadsheetApp.getActive();
 const defs={
  AppModules:["ModuleKey","ModuleName","SheetName","Group","AdminVisible","MainAppEligible","Confidential","Icon","SortOrder","Status"],
  PaymentGateways:["GatewayID","GatewayName","Provider","Mode","KeyId","KeySecretEncrypted","WebhookSecretEncrypted","WebhookURL","Currency","AutoReceiptAfterCapture","Status","Notes"],
  PaymentTransactions:["PaymentID","GatewayID","OrderID","ProviderPaymentID","DonationID","ReceiptNo","DonorName","Mobile","Amount","Currency","Status","Method","CreatedAt","CapturedAt","WebhookVerified","RawReference","Notes"]
 };
 Object.keys(defs).forEach(n=>ensure_(ss,n,defs[n]));
 ensure_(ss,"Designations",["DesignationID","DesignationName","Level","Rank","DepartmentRequired","CanHoldOffice","VolunteerLimit","Status","Notes"]);
 ensure_(ss,"MainAppAccess",["DesignationID","ModuleKey","View","Add","Edit","Delete","Approve","Scope","Status"]);
 ensure_(ss,"MemberLogins",["LoginID","MemberID","Phone","PasswordHash","MustChangePassword","FailedAttempts","LockedUntil","LastLogin","Status","PasswordChangedAt"]);
 ensure_(ss,"MemberStatusHistory",["HistoryID","MemberID","OldStatus","NewStatus","Reason","ChangedBy","ChangedAt"]);
 seedDesignationsV2_(ss); seedModulesV2_(ss); seedPayment_(ss); seedTempNational_(ss); seedAdmin_(ss); seedAccessV2_(ss);
 SpreadsheetApp.flush();return "Sewangan Android v2 repair completed";
}
function ensure_(ss,n,h){let s=ss.getSheetByName(n);if(!s)s=ss.insertSheet(n);let old=s.getLastColumn()?s.getRange(1,1,1,s.getLastColumn()).getValues()[0]:[];if(!old.filter(String).length)s.getRange(1,1,1,h.length).setValues([h]);else{let miss=h.filter(x=>!old.includes(x));if(miss.length)s.getRange(1,s.getLastColumn()+1,1,miss.length).setValues([miss])}s.getRange(1,1,1,s.getLastColumn()).setBackground("#0B6B36").setFontColor("#fff").setFontWeight("bold");s.setFrozenRows(1)}
function seedDesignationsV2_(ss){
 const s=ss.getSheetByName("Designations");s.getRange(2,1,Math.max(1,s.getMaxRows()-1),9).clearContent();
 const base=[
 ["DES-NAT-PRES","National President","National"],["DES-NAT-VP","Vice President","National"],["DES-CEO","CEO","National"],["DES-NAT-SEC","National Secretary","National"],["DES-NAT-TREAS","National Treasurer","National"],["DES-NAT-JS","Joint Secretary","National"],["DES-NAT-EXEC","National Executive Committee Member","National"]];
 const md=["Education","Health & Medical","Women Empowerment","Youth Affairs","Child Welfare","Social Justice","Rural Development","Employment & Skill Development","Environment Protection","Disaster Relief","Legal Aid","Media & Public Relations","Information Technology","Cultural & Sports","Membership Development","Finance & Donation Management","Research & Planning","Human Rights","Senior Citizen Welfare","Minority & Backward Welfare"];
 md.forEach((d,i)=>base.push(["DES-MD-"+String(i+1).padStart(2,"0"),"M.D. (Dept. of "+d+")","National",d]));
 base.push(["DES-STATE-PRES","State President","State"],["DES-STATE-SEC","State Secretary","State"],["DES-STATE-TREAS","State Treasurer","State"]);
 ["Organization & Membership","Finance & Donation","Program & Social Welfare","Media & Public Relations","Youth & Volunteer Management"].forEach((d,i)=>base.push(["DES-RM-"+String(i+1).padStart(2,"0"),"R.M. (Dept. of "+d+")","State",d]));
 base.push(["DES-DIST-PRES","District President","District"],["DES-DIST-SEC","District Secretary","District"],["DES-DIST-CO","District Coordinator","District"],["DES-BLOCK-PRES","Block President","Block"],["DES-BLOCK-CO","Block Coordinator","Block"],["DES-AM","Associate Member","Panchayat"],["DES-VOL","Volunteer","Panchayat"]);
 const rows=base.map((r,i)=>[r[0],r[1],r[2],i+1,!!r[3],r[2]!=="Panchayat",r[0]==="DES-AM"?20:0,"Active",r[3]||""]);s.getRange(2,1,rows.length,9).setValues(rows);
}
function seedModulesV2_(ss){
 const s=ss.getSheetByName("AppModules");s.getRange(2,1,Math.max(1,s.getMaxRows()-1),10).clearContent();
 const names=["Organization","Settings","Users","Roles","RoleAccess","MemberTypes","Members","MembershipApplications","Employees","EmployeeDocuments","TermsConditions","Volunteers","Beneficiaries","HelpRequests","CharityActivities","Donors","ReceiptSequence","Receipts","Donations","GoodsDonations","DonationCollection","OfflineCash","ReceiptBlocks","Fundraising","Programs","Projects","Events","Campaigns","Attendance","Leave","Accounts","Transactions","CashBook","BankBook","Ledger","Payments","Vouchers","Expenses","Inventory","InventoryMovements","Purchases","Vendors","Assets","Documents","Letters","DispatchRegister","Certificates","Meetings","Notifications","Numbering","AuditLog","Offices","Departments","Designations","Areas","MainAppAccess","Referrals","MemberStatusHistory","MemberLogins","MemberGeneratedDocuments","PaymentGateways","PaymentTransactions"];
 const confidential=["Settings","Users","Roles","RoleAccess","ReceiptSequence","ReceiptBlocks","Numbering","AuditLog","MainAppAccess","MemberLogins","PaymentGateways"];
 const rows=names.map((n,i)=>{const key=n.replace(/([a-z])([A-Z])/g,"$1_$2").toLowerCase();let g="Administration";if(["Members","MembershipApplications","MemberTypes","Volunteers","Referrals","MemberStatusHistory","MemberLogins","MemberGeneratedDocuments"].includes(n))g="Membership";else if(["Employees","EmployeeDocuments","TermsConditions","Attendance","Leave"].includes(n))g="Human Resources";else if(["Donors","Receipts","Donations","GoodsDonations","DonationCollection","OfflineCash","ReceiptBlocks","Fundraising","PaymentGateways","PaymentTransactions","ReceiptSequence"].includes(n))g="Donation & Fundraising";else if(["Accounts","Transactions","CashBook","BankBook","Ledger","Payments","Vouchers","Expenses"].includes(n))g="Finance";else if(["Inventory","InventoryMovements","Purchases","Vendors","Assets"].includes(n))g="Inventory & Assets";else if(["Programs","Projects","Events","Campaigns","Meetings"].includes(n))g="Programs";else if(["Beneficiaries","HelpRequests","CharityActivities"].includes(n))g="Welfare";else if(["Documents","Letters","DispatchRegister","Certificates","Notifications"].includes(n))g="Documents & Communication";else if(["Offices","Areas"].includes(n))g="Organization Network";return[key,n.replace(/([a-z])([A-Z])/g,"$1 $2"),n,g,true,!confidential.includes(n)&&n!=="Organization",confidential.includes(n),"grid",i+1,"Active"]});
 s.getRange(2,1,rows.length,10).setValues(rows);
}
function seedPayment_(ss){const s=ss.getSheetByName("PaymentGateways");if(s.getLastRow()===1)s.appendRow(["PG-RAZORPAY","Razorpay","Razorpay","Test","","","","","INR",true,"Inactive","Future-ready placeholder"])}
function sha_(x){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(x),Utilities.Charset.UTF_8).map(b=>("0"+(b&255).toString(16)).slice(-2)).join("")}
function seedTempNational_(ss){
 const ms=ss.getSheetByName("Members"),h=ms.getRange(1,1,1,ms.getLastColumn()).getValues()[0],id="SCT/2026/TEST/00001";
 let obj={MemberID:id,Name:"National Test",FirstName:"national",DOB:"1990-10-02",Mobile:"9999999999",Email:"test@sewangan.in",Address:"Delhi",State:"Delhi",OfficeID:"OFF-NAT-DELHI",DesignationID:"DES-NAT-SEC",Designation:"National Secretary",DepartmentID:"DEP-001",Department:"General",JoinDate:new Date(),ValidUntil:"2026-12-31",Status:"Active",StatusReason:"Temporary testing member",Notes:"TEMPORARY NATIONAL-LEVEL MAIN APP TEST ACCOUNT"};
 let v=ms.getDataRange().getValues(),row=0;for(let i=1;i<v.length;i++)if(String(v[i][0])===id)row=i+1;let data=h.map(k=>obj[k]??"");row?ms.getRange(row,1,1,h.length).setValues([data]):ms.appendRow(data);
 const ls=ss.getSheetByName("MemberLogins"),lh=ls.getRange(1,1,1,ls.getLastColumn()).getValues()[0],lo={LoginID:"9999999999",MemberID:id,Phone:"9999999999",PasswordHash:sha_("national02"),MustChangePassword:false,FailedAttempts:0,Status:"Active",PasswordChangedAt:new Date()};let lv=ls.getDataRange().getValues(),lr=0;for(let i=1;i<lv.length;i++)if(String(lv[i][1])===id)lr=i+1;let ld=lh.map(k=>lo[k]??"");lr?ls.getRange(lr,1,1,lh.length).setValues([ld]):ls.appendRow(ld);
}
function seedAdmin_(ss){const s=ss.getSheetByName("Users");if(s.getLastRow()===1)s.appendRow(["U001","Super Admin","admin@sewangan.in","CHANGE_ME_NOW","ROLE_SUPER","OFF-HO-SURSAND","","Active",""])}
function seedAccessV2_(ss){
 const s=ss.getSheetByName("MainAppAccess"),ds=ss.getSheetByName("Designations").getDataRange().getValues().slice(1),mods=ss.getSheetByName("AppModules").getDataRange().getValues().slice(1).filter(r=>r[5]===true);
 s.getRange(2,1,Math.max(1,s.getMaxRows()-1),9).clearContent();let out=[];
 ds.filter(r=>r[0]).forEach(d=>mods.forEach(m=>{let level=d[2],view=false,add=false,edit=false,del=false,approve=false,scope="Own";if(["National","State"].includes(level)){view=add=edit=approve=true;scope="Own + All Lower"}else if(level==="District"){view=add=edit=true;approve=["members","membership_applications","volunteers","donations","goods_donations","help_requests"].includes(m[0]);scope="Own + All Lower"}else if(level==="Block"){view=add=edit=true;scope="Own + All Lower"}else if(d[0]==="DES-AM"){view=["members","volunteers","donors","donations","goods_donations","programs","projects","events","campaigns","documents","certificates","meetings","notifications","beneficiaries","help_requests"].includes(m[0]);add=["volunteers","donors","donations","goods_donations","beneficiaries","help_requests"].includes(m[0]);edit=["volunteers","donors","beneficiaries","help_requests"].includes(m[0]);scope="Own + Direct Lower"}else{view=["programs","projects","events","campaigns","notifications","documents","certificates","meetings","attendance"].includes(m[0])}out.push([d[0],m[0],view,add,edit,del,approve,scope,"Active"])}));
 if(out.length)s.getRange(2,1,out.length,9).setValues(out);
}
