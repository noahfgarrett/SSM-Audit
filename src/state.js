export const S={
  screen:'upload',
  homeMode:'audit',
  ui:{navCollapsed:false},
  rules:{search:'',source:'all',category:'all'},
  session:null,
  comparison:null,
  updateInfo:null,
  updateModalDismissed:false,
};

export function resetSession(){
  S.session={name:'',snapshot:null,result:null,error:'',auditedAt:0,dashFilter:null,hiddenSources:[],hiddenSeverities:[],hiddenCategories:[],hiddenRules:[],search:'',sort:'severity-desc',groupBy:'none',collapsedGroups:[],cursor:-1,filterOpen:false,selectedFindingId:'',scrollTop:0,fullscreen:false,filteredCacheKey:'',filteredCacheRows:null,displayCacheKey:'',displayCacheRows:null,rowIndex:null,headerIdSet:null,hierarchy:null,hierarchySearch:'',hierarchyBuilding:'all',hierarchyDiscipline:'all',hierarchySystem:'all',hierarchyFindingsOnly:false,hierarchyExpandedKeys:[],hierarchyInitialized:false,hierarchyScrollTop:0,hierarchyFullscreen:false,hierarchyCacheKey:'',hierarchyCacheRows:null,hierarchyFocusKey:'',selectedHierarchyNodeKey:''};
}

export function resetComparison(){
  S.comparison={targetName:'',targetSnapshot:null,referenceName:'',referenceSnapshot:null,result:null,targetError:'',referenceError:'',selectedUpn:'',systemSearch:'',systemFilter:'different',systemSort:'upn-asc',detailTab:'hierarchy',rowSearch:'',rowFilter:'different',pairScrollTop:0,treeScrollTop:0,treeExpandedByUpn:{},fullscreen:false};
}

resetSession();
resetComparison();
