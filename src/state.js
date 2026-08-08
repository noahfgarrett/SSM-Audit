export const S={
  screen:'upload',
  session:null,
  updateInfo:null,
  updateModalDismissed:false,
};

export function resetSession(){
  S.session={name:'',snapshot:null,result:null,error:'',severity:'all',category:'all',search:'',sort:'severity',selectedFindingId:'',scrollTop:0,fullscreen:false};
}

resetSession();
