function apiDeleteCycle(adminCode, cycleKey){
  const sh = SpreadsheetApp.getActive().getSheetByName('CYCLE_MASTER');
  if(!sh) return {ok:false,error:'NO_SHEET'};
  const data = sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0])===String(cycleKey)){
      sh.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false,error:'NOT_FOUND'};
}
