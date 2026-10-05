
const SS=SpreadsheetApp.getActive();
const CACHE=CacheService.getScriptCache();
const SESSION_SECONDS=43200;
const MODULE_SHEETS={
  offices:"Offices",members:"Members",volunteers:"Volunteers",donations:"Donations",goodsdonations:"GoodsDonations",
  donors:"Donors",programs:"Programs",projects:"Projects",events:"Events",campaigns:"Campaigns",beneficiaries:"Beneficiaries",
  helprequests:"HelpRequests",charityactivities:"CharityActivities",attendance:"Attendance",documents:"Documents",
  certificates:"Certificates",meetings:"Meetings",notifications:"Notifications",membershipapplications:"MembershipApplications"
};
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function sheet_(n){const s=SS.getSheetByName(n);if(!s)throw Error("Missing sheet: "+n);return s}
function rows_(n){const v=sheet_(n).getDataRange().getDisplayValues();if(v.length<2)return[];const h=v[0];return v.slice(1).filter(r=>r.some(Boolean)).map(r=>Object.fromEntries(h.map((x,i)=>[x,r[i]])))}
function append_(n,o){const s=sheet_(n),h=s.getRange(1,1,1,s.getLastColumn()).getValues()[0];s.appendRow(h.map(k=>o[k]??""));return o}
function save_(n,o){const s=sheet_(n),v=s.getDataRange().getValues(),h=v[0],id=String(o[h[0]]||"");if(!id)throw Error(h[0]+" required");let rr=0;for(let i=1;i<v.length;i++)if(String(v[i][0])===id){rr=i+1;break}const row=h.map(k=>o[k]??"");rr?s.getRange(rr,1,1,h.length).setValues([row]):s.appendRow(row);return o}
function kv_(n,key="SettingKey",val="SettingValue"){return Object.fromEntries(rows_(n).map(r=>[r[key]||r.Key,r[val]??r.Value]))}
function uid_(p){return p+"-"+Utilities.getUuid().slice(0,8).toUpperCase()}
function audit_(user,action,module,id,details){append_("AuditLog",{LogID:uid_("LOG"),DateTime:new Date(),UserID:user||"SYSTEM",Action:action,Module:module,RecordID:id||"",Details:details||""})}
function hash_(password){const b=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,password,Utilities.Charset.UTF_8);return b.map(x=>("0"+(x&255).toString(16)).slice(-2)).join("")}
function firstName_(name){return String(name||"").trim().split(/\s+/)[0].toLowerCase().replace(/[^a-z]/g,"")}
function dobDay_(dob){const d=new Date(dob);if(isNaN(d))return"01";return String(d.getDate()).padStart(2,"0")}
function defaultPassword_(name,dob){return firstName_(name)+dobDay_(dob)}
function memberByPhone_(phone){return rows_("Members").find(x=>String(x.Mobile).replace(/\D/g,"")===String(phone).replace(/\D/g,""))}
function loginRow_(memberId){return rows_("MemberLogins").find(x=>x.MemberID===memberId)}
function memberLogin_(p){
  const m=memberByPhone_(p.phone); if(!m)throw Error("Invalid phone number or password");
  if(m.Status==="Blacklisted")throw Error("This member is blacklisted and cannot login.");
  if(m.Status!=="Active")throw Error("This member account is inactive.");
  const l=loginRow_(m.MemberID);if(!l||l.Status!=="Active"||l.PasswordHash!==hash_(p.password))throw Error("Invalid phone number or password");
  const token=Utilities.getUuid(),data={kind:"member",member:m,access:mainAccess_(m.DesignationID)};
  CACHE.put("S_"+token,JSON.stringify(data),SESSION_SECONDS);return{token,...data};
}
function adminLogin_(p){
  const u=rows_("Users").find(x=>(x.Email===p.user||x.UserID===p.user)&&x.Status!=="Inactive");
  if(!u)throw Error("Invalid login");
  const ok=String(u.PasswordHash)===String(p.password)||String(u.PasswordHash)===hash_(p.password);
  if(!ok)throw Error("Invalid login");
  const data={kind:"admin",user:u};const token=Utilities.getUuid();CACHE.put("S_"+token,JSON.stringify(data),SESSION_SECONDS);return{token,...data};
}
function session_(t){const x=CACHE.get("S_"+t);if(!x)throw Error("Session expired");return JSON.parse(x)}
function mainAccess_(designationId){return rows_("MainAppAccess").filter(x=>x.DesignationID===designationId&&x.Status!=="Inactive")}
function can_(s,module,perm){
  if(s.kind==="admin")return true;
  const a=(s.access||[]).find(x=>x.ModuleKey===module);
  if(!a||String(a.View).toLowerCase()!=="true")throw Error("Module access denied");
  if(perm&&String(a[perm]).toLowerCase()!=="true")throw Error(perm+" access denied");
  return a;
}
function descendants_(officeId){
  const all=rows_("Offices"),set=new Set([officeId]);let changed=true;
  while(changed){changed=false;all.forEach(o=>{if(o.ParentOfficeID&&set.has(o.ParentOfficeID)&&!set.has(o.OfficeID)){set.add(o.OfficeID);changed=true}})}
  return set;
}
function scoped_(s,module,data){
  if(s.kind==="admin")return data;
  const a=can_(s,module,"View"),m=s.member,scope=a.Scope||"Own";
  if(scope==="Own")return data.filter(r=>r.MemberID===m.MemberID||r.ReferredByMemberID===m.MemberID||r.AssociateMemberID===m.MemberID||r.CollectorUserID===m.MemberID);
  if(scope==="Own + Direct Lower")return data.filter(r=>r.MemberID===m.MemberID||r.ReferredByMemberID===m.MemberID||r.AssociateMemberID===m.MemberID||r.OfficeID===m.OfficeID||r.Panchayat===m.Panchayat);
  if(scope==="Own + All Lower"){const ids=descendants_(m.OfficeID);return data.filter(r=>!r.OfficeID||ids.has(r.OfficeID)||r.MemberID===m.MemberID)}
  return data;
}
function nextReceipt_(){const l=LockService.getScriptLock();l.waitLock(30000);try{const s=sheet_("ReceiptSequence"),r=s.getRange(2,1,1,6).getValues()[0],n=Number(r[2]||1),d=Number(r[3]||6);if(n>999999)throw Error("Receipt sequence exhausted for financial year");const no=(r[1]||"SDR")+"/"+(r[0]||"2627")+"/"+String(n).padStart(d,"0");s.getRange(2,3).setValue(n+1);s.getRange(2,5,1,2).setValues([[no,new Date()]]);return no}finally{l.releaseLock()}}
function nextMemberId_(){const y=Utilities.formatDate(new Date(),"Asia/Kolkata","yyyy");const n=rows_("Members").length+1;return"SCT/"+y+"/MB/"+String(n).padStart(5,"0")}
function approveMember_(s,p){
  const a=rows_("MembershipApplications").find(x=>x.ApplicationID===p.applicationId);if(!a)throw Error("Application not found");
  if(a.ApplicationStatus==="Approved")throw Error("Already approved");
  const memberId=nextMemberId_(),m={MemberID:memberId,Name:a.Name,FirstName:a.FirstName||firstName_(a.Name),FatherSpouse:a.FatherSpouse,DOB:a.DOB,Gender:a.Gender,Mobile:a.Mobile,Email:a.Email,Address:a.Address,State:a.State,District:a.District,Block:a.Block,Panchayat:a.Panchayat,AreaID:a.AreaID,OfficeID:a.OfficeID,DesignationID:a.DesignationID,Designation:a.MemberType,DepartmentID:a.DepartmentID,ReferredByMemberID:a.ReferredByMemberID,JoinDate:new Date(),Status:"Active"};
  save_("Members",m);
  const pw=defaultPassword_(m.Name,m.DOB);
  append_("MemberLogins",{LoginID:m.Mobile,MemberID:memberId,Phone:m.Mobile,PasswordHash:hash_(pw),MustChangePassword:true,FailedAttempts:0,Status:"Active",PasswordChangedAt:new Date()});
  a.ApplicationStatus="Approved";a.ReviewedBy=s.user.UserID;a.ReviewedAt=new Date();a.MemberID=memberId;save_("MembershipApplications",a);
  if(a.ReferredByMemberID)append_("Referrals",{ReferralID:uid_("REF"),ReferralType:"Membership",Date:new Date(),ReferrerMemberID:a.ReferredByMemberID,ReferrerName:a.ReferredByName,ReferredPersonID:memberId,ReferredPersonName:m.Name,RelatedRecordID:a.ApplicationID,Source:a.Source});
  audit_(s.user.UserID,"APPROVE","members",memberId,"New member approved");
  return{member:m,login:{phone:m.Mobile,defaultPassword:pw,mustChange:true}};
}
function setMemberStatus_(s,p){
  const m=rows_("Members").find(x=>x.MemberID===p.memberId);if(!m)throw Error("Member not found");
  const old=m.Status;m.Status=p.status;m.StatusReason=p.reason||"";if(p.status==="Blacklisted")m.BlacklistedDate=new Date();save_("Members",m);
  const l=loginRow_(m.MemberID);if(l){l.Status=p.status==="Active"?"Active":"Inactive";save_("MemberLogins",l)}
  append_("MemberStatusHistory",{HistoryID:uid_("MSH"),MemberID:m.MemberID,OldStatus:old,NewStatus:p.status,Reason:p.reason,ChangedBy:s.user.UserID,ChangedAt:new Date()});
  audit_(s.user.UserID,"STATUS","members",m.MemberID,old+" -> "+p.status);return m;
}
function resetPassword_(s,p){
  const m=rows_("Members").find(x=>x.MemberID===p.memberId);if(!m)throw Error("Member not found");
  let l=loginRow_(m.MemberID);if(!l)throw Error("Login not found");
  const pw=p.password||defaultPassword_(m.Name,m.DOB);l.PasswordHash=hash_(pw);l.MustChangePassword=true;l.Status=m.Status==="Active"?"Active":"Inactive";l.PasswordChangedAt=new Date();save_("MemberLogins",l);audit_(s.user.UserID,"RESET_PASSWORD","members",m.MemberID,"Password reset");return{phone:m.Mobile,newPassword:pw};
}
function changePassword_(s,p){
  if(s.kind!=="member")throw Error("Member login required");const l=loginRow_(s.member.MemberID);if(!l||l.PasswordHash!==hash_(p.oldPassword))throw Error("Current password is incorrect");if(String(p.newPassword).length<6)throw Error("New password must be at least 6 characters");l.PasswordHash=hash_(p.newPassword);l.MustChangePassword=false;l.PasswordChangedAt=new Date();save_("MemberLogins",l);return true;
}
function volunteerCount_(memberId){return rows_("Volunteers").filter(x=>x.MemberID===memberId||x.AssociateMemberID===memberId).length}
function addVolunteer_(s,p){
  if(s.kind!=="member")throw Error("Member login required");if(s.member.DesignationID!=="DES-AM")throw Error("Only Associate Member can add volunteers");
  if(volunteerCount_(s.member.MemberID)>=20)throw Error("Maximum 20 volunteers allowed");
  can_(s,"volunteers","Add");
  const id=uid_("VOL");append_("Volunteers",{VolunteerID:id,MemberID:"",Name:p.name,Mobile:p.mobile,Skills:p.skills,InterestArea:p.interest,Availability:p.availability,AssignedActivity:"",JoiningDate:new Date(),Status:"Pending",AssociateMemberID:s.member.MemberID,Panchayat:s.member.Panchayat});
  return{id};
}
function officeCards_(s){let a=rows_("Offices");if(s.kind==="member"){const acc=can_(s,"offices","View"),ids=descendants_(s.member.OfficeID);a=a.filter(x=>ids.has(x.OfficeID))}return a.map(o=>({...o,MemberCount:rows_("Members").filter(m=>m.OfficeID===o.OfficeID&&m.Status==="Active").length,LowerOfficeCount:rows_("Offices").filter(x=>x.ParentOfficeID===o.OfficeID).length}))}
function createDonation_(s,p){
  can_(s,"donations","Add");const no=nextReceipt_(),id=uid_("DON"),ref=s.kind==="member"?s.member.MemberID:"";
  append_("Donations",{DonationID:id,ReceiptNo:no,Date:new Date(),DonorID:p.donorId||"",DonorName:p.donorName,Amount:p.amount,PaymentMode:p.paymentMode,TransactionNo:p.transactionNo,Purpose:p.purpose,ReceivedAt:p.receivedAt||"",CollectorUserID:ref,ReferredByMemberID:ref,ReferredByName:s.kind==="member"?s.member.Name:"",Source:s.kind==="member"?"Member App":"Admin App",SyncStatus:"Synced",VerificationStatus:s.kind==="member"?"Pending":"Verified",Status:s.kind==="member"?"Pending Verification":"Received"});
  if(ref)append_("Referrals",{ReferralID:uid_("REF"),ReferralType:"Donation",Date:new Date(),ReferrerMemberID:ref,ReferrerName:s.member.Name,ReferredPersonID:p.donorId||"",ReferredPersonName:p.donorName,RelatedRecordID:id,Amount:p.amount,Source:"Member App"});
  return{donationId:id,receiptNo:no};
}
function doGet(){return json_({ok:true,app:"Sewangan ERP API",version:"final-revised"})}
function doPost(e){try{
  const q=JSON.parse(e.postData.contents||"{}"),a=q.action,p=q.payload||{};let s=null;
  if(!["memberLogin","adminLogin","publicSettings","joinRequest"].includes(a))s=session_(q.token);
  let d;
  if(a==="memberLogin")d=memberLogin_(p);
  else if(a==="adminLogin")d=adminLogin_(p);
  else if(a==="session")d=s;
  else if(a==="publicSettings")d={organization:kv_("Organization","Key","Value"),settings:kv_("Settings"),designations:rows_("Designations")};
  else if(a==="joinRequest"){const id=uid_("APP");append_("MembershipApplications",{ApplicationID:id,Date:new Date(),Source:p.source||"Website",MemberType:p.designation,DesignationID:p.designationId,Name:p.name,FirstName:firstName_(p.name),FatherSpouse:p.fatherSpouse,DOB:p.dob,Gender:p.gender,Mobile:p.mobile,Email:p.email,Address:p.address,State:p.state,District:p.district,Block:p.block,Panchayat:p.panchayat,AreaID:p.areaId,OfficeID:p.officeId,DepartmentID:p.departmentId,ReferredByMemberID:p.referredByMemberId||"",ReferredByName:p.referredByName||"",ApplicationStatus:"Pending"});d={applicationId:id}}
  else if(a==="mainAccess")d=s.kind==="member"?s.access:rows_("MainAppAccess");
  else if(a==="officeCards")d=officeCards_(s);
  else if(a==="list"){can_(s,p.module,"View");d=scoped_(s,p.module,rows_(MODULE_SHEETS[p.module]))}
  else if(a==="save"){can_(s,p.module,p.isEdit?"Edit":"Add");d=save_(MODULE_SHEETS[p.module],p.record)}
  else if(a==="addVolunteer")d=addVolunteer_(s,p);
  else if(a==="createDonation")d=createDonation_(s,p);
  else if(a==="approveMember"){if(s.kind!=="admin")throw Error("Admin required");d=approveMember_(s,p)}
  else if(a==="setMemberStatus"){if(s.kind!=="admin")throw Error("Admin required");d=setMemberStatus_(s,p)}
  else if(a==="resetPassword"){if(s.kind!=="admin")throw Error("Admin required");d=resetPassword_(s,p)}
  else if(a==="changePassword")d=changePassword_(s,p);
  else if(a==="memberDetails"){if(s.kind!=="admin")throw Error("Admin required");const m=rows_("Members").find(x=>x.MemberID===p.memberId);d={member:m,login:loginRow_(p.memberId),referrals:rows_("Referrals").filter(x=>x.ReferrerMemberID===p.memberId),documents:rows_("MemberGeneratedDocuments").filter(x=>x.MemberID===p.memberId)}}
  else if(a==="saveMainAccess"){if(s.kind!=="admin")throw Error("Admin required");p.rows.forEach(r=>save_("MainAppAccess",r));d=true}
  else throw Error("Unknown action");
  return json_({ok:true,data:d});
}catch(err){return json_({ok:false,error:String(err.message||err)})}}
