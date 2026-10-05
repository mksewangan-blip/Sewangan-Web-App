
/**
 * Sewangan ERP - Spreadsheet Repair / Upgrade
 * Run repairSewanganSpreadsheet() ONCE from the bound Apps Script project.
 * It preserves existing rows where possible, repairs/adds required sheets and headers,
 * seeds hierarchy masters, and creates the 6-digit receipt sequence.
 */
function repairSewanganSpreadsheet() {
  const ss = SpreadsheetApp.getActive();
  const defs = {
    Offices:["OfficeID","OfficeName","OfficeType","ParentOfficeID","State","District","Block","Address","PIN","Phone","Email","PresidentMemberID","InChargeMemberID","OpeningDate","Status","Notes"],
    Departments:["DepartmentID","DepartmentName","Level","ParentDepartmentID","Status","Notes"],
    Designations:["DesignationID","DesignationName","Level","Rank","DepartmentRequired","CanHoldOffice","VolunteerLimit","Status","Notes"],
    Areas:["AreaID","AreaType","State","District","Block","Panchayat","ParentAreaID","AssociateMemberID","Status","Notes"],
    MainAppAccess:["DesignationID","ModuleKey","View","Add","Edit","Delete","Approve","Scope","Status"],
    Referrals:["ReferralID","ReferralType","Date","ReferrerMemberID","ReferrerName","ReferredPersonID","ReferredPersonName","RelatedRecordID","Amount","Source","Notes"],
    MemberStatusHistory:["HistoryID","MemberID","OldStatus","NewStatus","Reason","ChangedBy","ChangedAt"],
    MemberLogins:["LoginID","MemberID","Phone","PasswordHash","MustChangePassword","FailedAttempts","LockedUntil","LastLogin","Status","PasswordChangedAt"],
    MemberGeneratedDocuments:["DocumentID","MemberID","DocumentType","DocumentNo","GeneratedAt","GeneratedBy","FileURL","Status"],
    Members:["MemberID","Name","FirstName","FatherSpouse","DOB","Gender","Mobile","Email","Address","State","District","Block","Panchayat","AreaID","OfficeID","DesignationID","Designation","DepartmentID","Department","ReferredByMemberID","JoinDate","ValidUntil","Fee","PaymentStatus","Status","StatusReason","BlacklistedDate","PhotoURL","Documents","Notes"],
    MembershipApplications:["ApplicationID","Date","Source","MemberType","DesignationID","Name","FirstName","FatherSpouse","DOB","Gender","Mobile","Email","Address","State","District","Block","Panchayat","AreaID","OfficeID","DepartmentID","ReferredByMemberID","ReferredByName","Amount","PaymentMode","TransactionNo","PaymentProofURL","PaymentStatus","ApplicationStatus","ReviewedBy","ReviewedAt","MemberID","Notes"],
    Donations:["DonationID","ReceiptNo","Date","DonorID","DonorName","Amount","PaymentMode","TransactionNo","PaymentProofURL","Purpose","Campaign","ReceivedAt","CollectorUserID","ReferredByMemberID","ReferredByName","Source","SyncStatus","VerificationStatus","Status"]
  };

  Object.keys(defs).forEach(name => repairSheet_(ss,name,defs[name]));

  const office = ss.getSheetByName("Offices");
  if (office.getLastRow() === 1) {
    office.getRange(2,1,2,16).setValues([
      ["OFF-HO-SURSAND","Head Office - Sursand","Head Office","","Bihar","Sitamarhi","Sursand","Sursand, Sitamarhi, Bihar","","","","","","","Active","Central registered/administrative head office"],
      ["OFF-NAT-DELHI","National Office - Delhi","National Office","","Delhi","","","Delhi","","","","","","","Active","National operational office"]
    ]);
  }

  seedMaster_(ss,"Departments",[
    ["DEP-001","General","All","","Active",""],
    ["DEP-002","Organization & Membership","All","","Active",""],
    ["DEP-003","Finance & Donation Management","All","","Active",""],
    ["DEP-004","Program & Social Welfare","All","","Active",""],
    ["DEP-005","Media & PR","All","","Active",""],
    ["DEP-006","Youth & Volunteer Management","All","","Active",""],
    ["DEP-007","Education","All","","Active",""],
    ["DEP-008","Health & Medical","All","","Active",""],
    ["DEP-009","Women Empowerment","All","","Active",""],
    ["DEP-010","IT","All","","Active",""]
  ]);

  seedMaster_(ss,"Designations",[
    ["DES-VOL","Volunteer","Panchayat",100,false,false,0,"Active","Volunteer under Associate Member"],
    ["DES-AM","Associate Member","Panchayat",90,false,false,20,"Active","One per Panchayat; max 20 volunteers"],
    ["DES-BLOCK-CO","Block Coordinator","Block",80,false,true,0,"Active",""],
    ["DES-BLOCK-PRES","Block President","Block",79,false,true,0,"Active",""],
    ["DES-DIST-CO","District Coordinator","District",70,false,true,0,"Active",""],
    ["DES-DIST-PRES","District President","District",68,false,true,0,"Active",""],
    ["DES-STATE-RM","Regional Manager","State",60,true,true,0,"Active",""],
    ["DES-STATE-PRES","State President","State",57,false,true,0,"Active",""],
    ["DES-NAT-MD","Managing Director","National",40,true,true,0,"Active",""],
    ["DES-NAT-SEC","National Secretary","National",37,false,true,0,"Active",""],
    ["DES-CEO","CEO","National",36,false,true,0,"Active",""],
    ["DES-NAT-PRES","National President","National",34,false,true,0,"Active",""]
  ]);

  repairReceiptSequence_(ss);
  seedMainAppAccess_(ss);

  const members=ss.getSheetByName("Members");
  const statusCol=defs.Members.indexOf("Status")+1;
  members.getRange(2,statusCol,Math.max(1000,members.getMaxRows()-1),1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["Active","Inactive","Blacklisted"],true).build());

  SpreadsheetApp.flush();
  return "Sewangan spreadsheet repaired/upgraded successfully.";
}

function repairSheet_(ss,name,headers){
  let s=ss.getSheetByName(name);
  if(!s) s=ss.insertSheet(name);
  const old=s.getLastColumn()?s.getRange(1,1,1,s.getLastColumn()).getValues()[0]:[];
  const missing=headers.filter(h=>!old.includes(h));
  if(s.getLastRow()===0 || old.filter(String).length===0){
    s.getRange(1,1,1,headers.length).setValues([headers]);
  } else if(missing.length){
    s.getRange(1,s.getLastColumn()+1,1,missing.length).setValues([missing]);
  }
  const lc=s.getLastColumn();
  s.getRange(1,1,1,lc).setBackground("#0B6B36").setFontColor("#FFFFFF").setFontWeight("bold").setWrap(true);
  s.setFrozenRows(1);
}

function seedMaster_(ss,name,rows){
  const s=ss.getSheetByName(name);
  if(s.getLastRow()===1 && rows.length) s.getRange(2,1,rows.length,rows[0].length).setValues(rows);
}

function repairReceiptSequence_(ss){
  let s=ss.getSheetByName("ReceiptSequence");
  if(!s){s=ss.insertSheet("ReceiptSequence");s.appendRow(["FinancialYear","Prefix","NextNumber","Digits","LastIssuedNo","UpdatedAt"]);}
  if(s.getLastRow()===1)s.appendRow(["2627","SDR",1,6,"",""]);
  else s.getRange(2,4).setValue(6);
  let n=ss.getSheetByName("Numbering");
  if(n){
    const v=n.getDataRange().getValues();
    for(let i=1;i<v.length;i++){
      if(String(v[i][0]).toLowerCase().includes("receipt")){
        n.getRange(i+1,5).setValue(6);
        n.getRange(i+1,6).setValue("SDR/2627/000001");
      }
    }
  }
}

function seedMainAppAccess_(ss){
  const s=ss.getSheetByName("MainAppAccess");
  if(s.getLastRow()>1)return;
  const ds=ss.getSheetByName("Designations").getDataRange().getValues().slice(1);
  const mods=["dashboard","offices","members","volunteers","donations","goodsdonations","donors","programs","projects","events","campaigns","beneficiaries","helprequests","charityactivities","attendance","documents","certificates","meetings","notifications","reports"];
  const rows=[];
  ds.filter(r=>r[0]).forEach(d=>{
    mods.forEach(m=>{
      const level=d[2], am=d[0]==="DES-AM";
      let view=["dashboard","programs","projects","events","campaigns","notifications"].includes(m);
      let add=false,edit=false,del=false,approve=false,scope="Own";
      if(am){view=view||["volunteers","donations","goodsdonations","donors","documents","certificates","meetings"].includes(m);add=["volunteers","donations","goodsdonations","donors"].includes(m);edit=m==="volunteers";scope="Own + Direct Lower";}
      if(["Block","District","State","National"].includes(level)){view=true;add=m!=="reports";edit=m!=="reports";scope="Own + All Lower";}
      if(["District","State","National"].includes(level))approve=["members","volunteers","donations","goodsdonations","helprequests"].includes(m);
      rows.push([d[0],m,view,add,edit,del,approve,scope,"Active"]);
    });
  });
  if(rows.length)s.getRange(2,1,rows.length,9).setValues(rows);
}
