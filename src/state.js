export const S={
  screen:'upload',
  session:null,
  updateInfo:null,
  updateModalDismissed:false,
};

export function resetSession(){
  S.session={name:'',snapshot:null,result:null,error:'',hiddenSources:[],hiddenSeverities:[],hiddenCategories:[],hiddenRules:[],search:'',sort:'severity-desc',filterOpen:false,selectedFindingId:'',scrollTop:0,fullscreen:false,filteredCacheKey:'',filteredCacheRows:null};
}

resetSession();
