function createBinder() {
  
var spread_sheet = SpreadsheetApp.getActiveSpreadsheet();
//set response sheet as active
var response_sheet = spread_sheet.getSheetByName('Form Responses 1');


//count number of responses
var response_vals= response_sheet.getDataRange().getValues()
var i = 0;
var ebo_idxs = [];
var dir_idxs = [];
var pc_idxs = [];
var active_idxs = [];
var sas_list = [];
var loa_list = [];

try{
  eboard_sheet = spread_sheet.getSheetByName('Eboard');
  spread_sheet.deleteSheet(eboard_sheet);
  director_sheet = spread_sheet.getSheetByName('Directors');
  spread_sheet.deleteSheet(director_sheet);
  pc_sheet = spread_sheet.getSheetByName('Pledge Committee');
  spread_sheet.deleteSheet(pc_sheet);
  ab_sheet = spread_sheet.getSheetByName('Active Brothers');
  spread_sheet.deleteSheet(ab_sheet);
  sas_sheet = spread_sheet.getSheetByName('SAS');
  spread_sheet.deleteSheet(sas_sheet);
  loa_sheet = spread_sheet.getSheetByName('LOA');
  spread_sheet.deleteSheet(loa_sheet);

}
catch{
  x = 0;
}
for (i = response_vals.length; i > 0; i--)
{

  if (response_vals[i] != null){
  if(response_vals[i][4] == "Eboard")
  {
    if(checkDuplicates(response_vals[i], ebo_idxs)){
      if(response_vals[i][0] == "President"){
        ebo_idxs.splice(0,0,response_vals[i]);
      }
      else if (response_vals[i][0] == "Executive Vice President"){
        if(ebo_idxs[0][0] != "President"){
          ebo_idxs.splice(0,0,response_vals[i]);
        }
        else{
          ebo_idxs.splice(1,0,response_vals[i]);
        }
      }
      else{
        ebo_idxs.push(response_vals[i]);
      }
    }
    
  }
  else if(response_vals[i][4] == "Director")
  {
    if(checkDuplicates(response_vals[i], dir_idxs)){
    dir_idxs.push(response_vals[i])
    }
  }
  else if(response_vals[i][4] == "Pledge Committee")
  {
    if(checkDuplicates(response_vals[i], pc_idxs)){
    pc_idxs.push(response_vals[i])
    }
  }
  else if(response_vals[i][4] == "SAS"){
    if(checkDuplicates(response_vals[i], sas_list)){
    sas_list.push(response_vals[i])
    }
  }
  else if(response_vals[i][4] == "LOA"){
    if(checkDuplicates(response_vals[i], loa_list)){
    loa_list.push(response_vals[i])
    }
  }
  else{
    if(checkDuplicates(response_vals[i], active_idxs)){
    active_idxs.push(response_vals[i]);
    }
  }
}
}

delete response_vals;

//create new sheets
spread_sheet.insertSheet().setName("Eboard");
spread_sheet.insertSheet().setName("Directors");
spread_sheet.insertSheet().setName("Pledge Committee");
spread_sheet.insertSheet().setName("Active Brothers");
spread_sheet.insertSheet().setName("SAS");
spread_sheet.insertSheet().setName("LOA");

var ebo_sheet = spread_sheet.getSheetByName("Eboard");
var dir_sheet = spread_sheet.getSheetByName("Directors");
var pc_sheet = spread_sheet.getSheetByName("Pledge Committee");
var ac_sheet = spread_sheet.getSheetByName("Active Brothers");
var sas_sheet = spread_sheet.getSheetByName("SAS");
var loa_sheet = spread_sheet.getSheetByName("LOA");

//sort remaining entries by name
var act_len = active_idxs.length;
var sorted_actives = [];

for(i = 0; i < act_len; i++){
  if (sorted_actives.length == 0){

    sorted_actives.splice(0,0,active_idxs[0]);
    inserted = true;
  }
  //find alpha first name and place it in the list where it belongs
  else{
    var inserted = false;
    for(j = 0; j < sorted_actives.length; j++){
      if(active_idxs[i][2] < sorted_actives[j][2])
      {
        sorted_actives.splice(j, 0,active_idxs[i]);
        inserted = true;
        break;
      }
      
    }
    if (inserted == false){
      sorted_actives.push(active_idxs[i]);
    }
  }

  }
delete active_idxs;

var big_list =   [ebo_idxs,  dir_idxs,  pc_idxs,  sorted_actives,sas_list,loa_list];
var sheet_list = [ebo_sheet, dir_sheet, pc_sheet, ac_sheet, sas_sheet, loa_sheet];
for (i = 0; i < big_list.length;i++){
  if (big_list[i] != null){
    var row = 0
    var length = big_list[i].length

    for(j = 0; j < length; j++){
      row++;
      row = createEntry(big_list[i][j], row, sheet_list[i]);
    }
  }
}


}
function createEntry(current_person,row,sheet){

  
  var position = current_person[0];
  var name = current_person[2];
  var email = current_person[3];
  var hometown = current_person[5];
  var year = current_person[6];
  var major_etc = current_person[7];
  var work = current_person[8];
  var camp_inv = current_person[9];
  var interests = current_person[10];
  var fav_mom = current_person[11];
  var head_link = current_person[12];

  var loop_array = [position, major_etc, work, camp_inv, year, hometown, interests, fav_mom];
  var picture_cell = sheet.getRange(row,1);
  var name_cell = sheet.getRange(row,2);
  
  
  sheet.getRange(row,1,6).mergeVertically();//merge column A rows 0-5

  // Insert the image in cell A1.

  picture_cell.setValue("picture");
  
  sheet.setColumnWidths(2,6, 200);
  sheet.getRange(row,2,1,6 ).merge();
  name_cell.setValue(name).setBackground("#5581FF").setHorizontalAlignment("center").setFontWeight("bold");

  var cell_titles = ["Current Position in Chapter", "Major(s), Minor(s), Certificate(s)", "Work/AKPsi Experience", "Campus Involvement(s)", "Year", "Hometown", "Interests and Hobbies", "Favorite AKPsi Moment"];

  var loop_row = 0;
  var loop_column = 2;
  for(var i = 0; i < cell_titles.length; i++){

    title_cell = sheet.getRange(((loop_row % 4) + 1 + row),loop_column);
    title_cell.setValue(cell_titles[i]).setBackground("#F7FF8E").setFontWeight("bold");
    entry_cell = sheet.getRange(((loop_row % 4) + 1+ row),loop_column + 1);
    entry_cell.setValue(loop_array[i]).setBackground("#F7FF8E");
    sheet.getRange(((loop_row % 4 + 1) + row),loop_column + 1,1,2 ).merge();

    loop_row ++;

    if(loop_row == 4){
      loop_column += 3;
    }
  } 

  var email_cell = sheet.getRange(row+ 5,2);
  sheet.getRange(row + 5,2,1,6 ).merge();
  email_cell.setValue(email).setHorizontalAlignment("center").setBackground("#F7FF8E").setFontWeight("bold");

  sheet.getRange(row,2,6,6 ).setWrap(true);
  return (row + 6);
}

function checkDuplicates(new_entry, current_array){
  for(i = 0; i < current_array.length; i++){
    if ((new_entry[2] == current_array[i][2]) || new_entry[3] == current_array[i]){
      return false;
    }
  }
  return true;
}
