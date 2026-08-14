export const S={
  screen:'upload',
  homeMode:'audit',
  rules:{search:'',source:'all',category:'all'},
  session:null,
  comparison:null,
  updateInfo:null,
  updateModalDismissed:false,
};

export function resetSession(){
  S.session={name:'',snapshot:null,result:null,error:'',hiddenSources:[],hiddenSeverities:[],hiddenCategories:[],hiddenRules:[],search:'',sort:'severity-desc',filterOpen:false,selectedFindingId:'',scrollTop:0,fullscreen:false,filteredCacheKey:'',filteredCacheRows:null,hierarchy:null,hierarchySearch:'',hierarchyBuilding:'all',hierarchyDiscipline:'all',hierarchySystem:'all',hierarchyExpandedKeys:[],hierarchyInitialized:false,hierarchyScrollTop:0,hierarchyFullscreen:false,hierarchyCacheKey:'',hierarchyCacheRows:null,selectedHierarchyNodeKey:''};
}

export function resetComparison(){
  S.comparison={targetName:'',targetSnapshot:null,referenceName:'',referenceSnapshot:null,result:null,targetError:'',referenceError:'',selectedUpn:'',systemSearch:'',systemFilter:'different',systemSort:'upn-asc',rowSearch:'',rowFilter:'different',pairScrollTop:0,treeScrollTop:0,treeExpandedByUpn:{},fullscreen:false};
}

resetSession();
resetComparison();
