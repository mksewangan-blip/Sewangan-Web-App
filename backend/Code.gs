
const SS=SpreadsheetApp.getActive();
const CACHE=CacheService.getScriptCache(), SESSION=PropertiesService.getScriptProperties();
const SESSION_TTL=21600, DATA_TTL=120;
function out_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function sh_(n){const s=SS.getSheetByName(n);if(!s)throw Error("Sheet not found: "+n);return s}
function key_(n){return "D_"+n}
function clear_(n){CACHE.remove(key_(n))}
function table_(n,fresh){
  const k=key_(n); if(!fresh){const c=CACHE.get(k);if(c)return JSON.parse(c)}
  const s=sh_(n),v=s.getDataRange().getDisplayValues(); if(!v.length)return[];
  const h=v[0],r=v.slice(1).filter(x=>x.some(Boolean)).map(x=>Object.fromEntries(h.map((z,i)=>[z,x[i]])));
  try{CACHE.put(k,JSON.stringify(r),DATA_TTL)}catch(e){}
  return r;
}
function headers_(n){return sh_(n).getRange(1,1,1,sh_(n).getLastColumn()).getDisplayValues()[0]}
function append_(n,o){const s=sh_(n),h=headers_(n);s.appendRow(h.map(k=>o[k]??""));clear_(n);return o}
function save_(n,o){const s=sh_(n),h=headers_(n),v=s.getDataRange().getValues(),id=String(o[h[0]]||"");if(!id)throw Error(h[0]+" required");let rr=0;for(let i=1;i<v.length;i++)if(String(v[i][0])===id){rr=i+1;break}const row=h.map(k=>o[k]??"");rr?s.getRange(rr,1,1,h.length).setValues([row]):s.appendRow(row);clear_(n);return o}
function hash_(x){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(x),Utilities.Charset.UTF_8).map(b=>("0"+(b&255).toString(16)).slice(-2)).join("")}
function token_(data){const t=Utilities.getUuid();SESSION.setProperty("S_"+t,JSON.stringify({exp:Date.now()+SESSION_TTL*1000,...data}));return t}
function sess_(t){const x=SESSION.getProperty("S_"+t);if(!x)throw Error("Session expired");const d=JSON.parse(x);if(d.exp<Date.now()){SESSION.deleteProperty("S_"+t);throw Error("Session expired")}return d}
function modules_(){return table_("AppModules").filter(x=>x.Status!=="Inactive")}
function module_(key){const m=modules_().find(x=>x.ModuleKey===key);if(!m)throw Error("Unknown module: "+key);return m}
function access_(designation){return table_("MainAppAccess").filter(x=>x.DesignationID===designation&&x.Status!=="Inactive")}
function bool_(v){return String(v).toLowerCase()==="true"}
function memberByPhone_(p){const q=String(p).replace(/\D/g,"");return table_("Members").find(x=>String(x.Mobile).replace(/\D/g,"")===q)}
function loginRow_(id){return table_("MemberLogins").find(x=>x.MemberID===id)}
function memberLogin_(p){
  const m=memberByPhone_(p.phone); if(!m)throw Error("Invalid phone number or password");
  if(m.Status==="Blacklisted")throw Error("This member is blacklisted and cannot login.");
  if(m.Status!=="Active")throw Error("This member account is inactive.");
  const l=loginRow_(m.MemberID); if(!l||l.Status!=="Active"||l.PasswordHash!==hash_(p.password))throw Error("Invalid phone number or password");
  const a=access_(m.DesignationID),t=token_({kind:"member",memberId:m.MemberID,designationId:m.DesignationID});
  return{token:t,member:m,access:a,modules:modules_().filter(x=>bool_(x.MainAppEligible)&&a.some(y=>y.ModuleKey===x.ModuleKey&&bool_(y.View)))};
}
function adminLogin_(p){
  const u=table_("Users").find(x=>(x.UserID===p.user||x.Email===p.user)&&x.Status!=="Inactive");if(!u)throw Error("Invalid login");
  if(!(u.PasswordHash===p.password||u.PasswordHash===hash_(p.password)))throw Error("Invalid login");
  return{token:token_({kind:"admin",userId:u.UserID,roleId:u.RoleID}),user:u,modules:modules_().filter(x=>bool_(x.AdminVisible))};
}
function hydrate_(s){
  if(s.kind==="admin"){s.user=table_("Users").find(x=>x.UserID===s.userId);return s}
  s.member=table_("Members").find(x=>x.MemberID===s.memberId);if(!s.member||s.member.Status!=="Active")throw Error("Member access disabled");
  s.access=access_(s.designationId);return s;
}
function permission_(s,key,perm){
  if(s.kind==="admin")return{Scope:"All"};
  const a=s.access.find(x=>x.ModuleKey===key);if(!a||!bool_(a.View))throw Error("Access denied");
  if(perm&&perm!=="View"&&!bool_(a[perm]))throw Error(perm+" permission denied");return a;
}
function descendants_(officeId){const all=table_("Offices"),set=new Set([officeId]);let c=true;while(c){c=false;all.forEach(o=>{if(o.ParentOfficeID&&set.has(o.ParentOfficeID)&&!set.has(o.OfficeID)){set.add(o.OfficeID);c=true}})}return set}
function scope_(s,key,data){
  if(s.kind==="admin")return data;
  const a=permission_(s,key,"View"),m=s.member,scope=a.Scope||"Own";
  if(scope==="Own")return data.filter(r=>r.MemberID===m.MemberID||r.ReferredByMemberID===m.MemberID||r.AssociateMemberID===m.MemberID||r.CollectorUserID===m.MemberID);
  if(scope==="Own + Direct Lower")return data.filter(r=>r.MemberID===m.MemberID||r.ReferredByMemberID===m.MemberID||r.AssociateMemberID===m.MemberID||r.OfficeID===m.OfficeID||r.Panchayat===m.Panchayat);
  if(scope==="Own + All Lower"){const ids=descendants_(m.OfficeID);return data.filter(r=>!r.OfficeID||ids.has(r.OfficeID)||r.MemberID===m.MemberID)}
  return data;
}
function dashboard_(s){
 const names=["Members","Employees","Volunteers","Donors","Donations","Beneficiaries","Projects","Events"];
 const counts={};names.forEach(n=>{try{counts[n]=s.kind==="admin"?table_(n).length:scope_(s,n.toLowerCase(),table_(n)).length}catch(e){counts[n]=0}});
 return{counts,recent:table_("Notifications").slice(-5).reverse()};
}
function bootstrap_(s){
 const mods=s.kind==="admin"?modules_().filter(x=>bool_(x.AdminVisible)):modules_().filter(x=>bool_(x.MainAppEligible)&&s.access.some(a=>a.ModuleKey===x.ModuleKey&&bool_(a.View)));
 return{session:s.kind==="admin"?{kind:"admin",user:s.user}:{kind:"member",member:s.member},modules:mods,access:s.access||[],dashboard:dashboard_(s),offices:officeCards_(s)};
}
function officeCards_(s){
 let os=table_("Offices");if(s.kind==="member"){permission_(s,"offices","View");const ids=descendants_(s.member.OfficeID);os=os.filter(x=>ids.has(x.OfficeID))}
 const members=table_("Members"),all=table_("Offices");
 return os.map(o=>({...o,MemberCount:members.filter(m=>m.OfficeID===o.OfficeID&&m.Status==="Active").length,LowerOfficeCount:all.filter(x=>x.ParentOfficeID===o.OfficeID).length}));
}
function nextReceipt_(){const l=LockService.getScriptLock();l.waitLock(30000);try{const s=sh_("ReceiptSequence"),r=s.getRange(2,1,1,6).getValues()[0],n=Number(r[2]||1),d=Number(r[3]||6);const no=(r[1]||"SDR")+"/"+(r[0]||"2627")+"/"+String(n).padStart(d,"0");s.getRange(2,3).setValue(n+1);s.getRange(2,5,1,2).setValues([[no,new Date()]]);clear_("ReceiptSequence");return no}finally{l.releaseLock()}}
function uid_(p){return p+"-"+Utilities.getUuid().slice(0,8).toUpperCase()}
function approve_(s,p){
 const a=table_("MembershipApplications",true).find(x=>x.ApplicationID===p.applicationId);if(!a)throw Error("Application not found");
 const id="SCT/"+Utilities.formatDate(new Date(),"Asia/Kolkata","yyyy")+"/MB/"+String(table_("Members").length+1).padStart(5,"0");
 const first=String(a.Name||"").trim().split(/\s+/)[0].toLowerCase().replace(/[^a-z]/g,""),d=new Date(a.DOB),day=isNaN(d)?"01":String(d.getDate()).padStart(2,"0"),pw=first+day;
 const m={MemberID:id,Name:a.Name,FirstName:first,DOB:a.DOB,Mobile:a.Mobile,Email:a.Email,Address:a.Address,State:a.State,District:a.District,Block:a.Block,Panchayat:a.Panchayat,AreaID:a.AreaID,OfficeID:a.OfficeID,DesignationID:a.DesignationID,Designation:a.MemberType,DepartmentID:a.DepartmentID,ReferredByMemberID:a.ReferredByMemberID,JoinDate:new Date(),Status:"Active"};
 append_("Members",m);append_("MemberLogins",{LoginID:m.Mobile,MemberID:id,Phone:m.Mobile,PasswordHash:hash_(pw),MustChangePassword:true,FailedAttempts:0,Status:"Active",PasswordChangedAt:new Date()});
 a.ApplicationStatus="Approved";a.ReviewedBy=s.userId;a.ReviewedAt=new Date();a.MemberID=id;save_("MembershipApplications",a);return{member:m,phone:m.Mobile,defaultPassword:pw};
}
function setStatus_(s,p){let m=table_("Members",true).find(x=>x.MemberID===p.memberId);if(!m)throw Error("Member not found");const old=m.Status;m.Status=p.status;m.StatusReason=p.reason||"";m.BlacklistedDate=p.status==="Blacklisted"?new Date():"";save_("Members",m);let l=loginRow_(m.MemberID);if(l){l.Status=p.status==="Active"?"Active":"Inactive";save_("MemberLogins",l)}append_("MemberStatusHistory",{HistoryID:uid_("MSH"),MemberID:m.MemberID,OldStatus:old,NewStatus:p.status,Reason:p.reason||"",ChangedBy:s.userId,ChangedAt:new Date()});return m}
function resetPw_(s,p){const m=table_("Members").find(x=>x.MemberID===p.memberId),l=loginRow_(p.memberId);if(!m||!l)throw Error("Member/login not found");const d=new Date(m.DOB),pw=(String(m.FirstName||m.Name).split(/\s+/)[0].toLowerCase())+(isNaN(d)?"01":String(d.getDate()).padStart(2,"0"));l.PasswordHash=hash_(pw);l.MustChangePassword=true;l.PasswordChangedAt=new Date();save_("MemberLogins",l);return{phone:m.Mobile,password:pw}}
function changePw_(s,p){if(s.kind!=="member")throw Error("Member login required");const l=loginRow_(s.memberId);if(l.PasswordHash!==hash_(p.oldPassword))throw Error("Current password incorrect");if(String(p.newPassword).length<6)throw Error("Minimum 6 characters");l.PasswordHash=hash_(p.newPassword);l.MustChangePassword=false;l.PasswordChangedAt=new Date();save_("MemberLogins",l);return true}
function donation_(s,p){permission_(s,"donations","Add");const no=nextReceipt_(),id=uid_("DON"),mid=s.kind==="member"?s.memberId:"";append_("Donations",{DonationID:id,ReceiptNo:no,Date:new Date(),DonorID:p.DonorID||"",DonorName:p.DonorName,Amount:p.Amount,PaymentMode:p.PaymentMode,TransactionNo:p.TransactionNo||"",Purpose:p.Purpose||"General Donation",CollectorUserID:mid,ReferredByMemberID:mid,Source:s.kind==="member"?"Member App":"Admin App",SyncStatus:"Synced",VerificationStatus:s.kind==="member"?"Pending":"Verified",Status:s.kind==="member"?"Pending Verification":"Received"});return{DonationID:id,ReceiptNo:no}}
function doGet(){return out_({ok:true,app:"Sewangan ERP API",version:"android-final-v2"})}
function doPost(e){try{
 const q=JSON.parse(e.postData.contents||"{}"),a=q.action,p=q.payload||{};let s=null;if(!["memberLogin","adminLogin"].includes(a))s=hydrate_(sess_(q.token));
 let d;
 if(a==="memberLogin")d=memberLogin_(p);
 else if(a==="adminLogin")d=adminLogin_(p);
 else if(a==="bootstrap")d=bootstrap_(s);
 else if(a==="list"){const m=module_(p.module);permission_(s,p.module,"View");d=scope_(s,p.module,table_(m.SheetName))}
 else if(a==="save"){const m=module_(p.module);permission_(s,p.module,p.isEdit?"Edit":"Add");d=save_(m.SheetName,p.record)}
 else if(a==="officeCards")d=officeCards_(s);
 else if(a==="approveMember"){if(s.kind!=="admin")throw Error("Admin required");d=approve_(s,p)}
 else if(a==="setMemberStatus"){if(s.kind!=="admin")throw Error("Admin required");d=setStatus_(s,p)}
 else if(a==="resetPassword"){if(s.kind!=="admin")throw Error("Admin required");d=resetPw_(s,p)}
 else if(a==="changePassword")d=changePw_(s,p);
 else if(a==="createDonation")d=donation_(s,p);
 else if(a==="mainAccess"){if(s.kind!=="admin")throw Error("Admin required");d=table_("MainAppAccess")}
 else if(a==="saveMainAccess"){if(s.kind!=="admin")throw Error("Admin required");p.rows.forEach(r=>save_("MainAppAccess",r));d=true}
 else if(a==="memberDetails"){if(s.kind!=="admin")throw Error("Admin required");d={member:table_("Members").find(x=>x.MemberID===p.memberId),documents:table_("MemberGeneratedDocuments").filter(x=>x.MemberID===p.memberId),referrals:table_("Referrals").filter(x=>x.ReferrerMemberID===p.memberId)}}
 else throw Error("Unknown action");
 return out_({ok:true,data:d});
}catch(err){return out_({ok:false,error:String(err.message||err)})}}
