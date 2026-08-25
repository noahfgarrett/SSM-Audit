export const S={
  screen:'upload',
  homeMode:'audit',
  ui:{navCollapsed:false},
  rules:{search:'',source:'all',category:'all',disabled:[]},
  session:null,
  comparison:null,
  updateInfo:null,
  updateCheck:null,
  updateRetried:false,
  updateModalDismissed:false,
};

export function resetSession(){
  S.session={name:'',snapshot:null,result:null,error:'',auditedAt:0,status:null,dimFilters:{discipline:[],milestone:[],upn:[],building:[]},dimOptionCache:null,filterOptionCache:null,hiddenSources:[],hiddenSeverities:[],hiddenCategories:[],hiddenRules:[],search:'',sort:'severity-desc',groupBy:'none',collapsedGroups:[],cursor:-1,panelOpen:false,panelSearch:'',selectedFindingId:'',scrollTop:0,actioned:null,actionedRev:0,hideActioned:false,excluded:null,excludedRev:0,modifySearch:'',modifyOpenRules:[],modifyClosedCats:[],modifyCountMode:'count',modifyScrollTop:0,modifyGroupsKey:'',modifyGroupsVal:null,dashBreakdownOpen:false,fullscreen:false,filteredCacheKey:'',filteredCacheRows:null,displayCacheKey:'',displayCacheRows:null,scopedCacheKey:'',scopedCache:null,rowIndex:null,headerIdSet:null,hierarchy:null,hierarchySearch:'',hierarchyBuilding:'all',hierarchyDiscipline:'all',hierarchySystem:'all',hierarchyFindingsOnly:false,hierarchyExpandedKeys:[],hierarchyInitialized:false,hierarchyScrollTop:0,hierarchyFullscreen:false,hierarchyCacheKey:'',hierarchyCacheRows:null,hierarchyFocusKey:'',selectedHierarchyNodeKey:''};
}

export function resetComparison(){
  S.comparison={targetName:'',targetSnapshot:null,referenceName:'',referenceSnapshot:null,result:null,targetError:'',referenceError:'',selectedUpn:'',systemSearch:'',systemFilter:'different',systemSort:'upn-asc',detailTab:'hierarchy',rowSearch:'',rowFilter:'different',pairScrollTop:0,treeScrollTop:0,treeExpandedByUpn:{},fullscreen:false};
}

resetSession();
resetComparison();
