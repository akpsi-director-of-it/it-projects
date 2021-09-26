function createBinder() {

//Things that might/want to change semester to semester
//********************************************************************************************************************** */ 

var response_sheet_name = 'Form Responses 1';//Name of the response sheet, this is how the code finds responses
var sheet_name_list = ["Eboard","Directors","Pledge Committee", "Active Brothers", "SAS", "LOA"];//names of the different sheets that you want to add to the binder
var categories = ["Current Position in Chapter", "Major(s), Minor(s), Certificate(s)", "Work/AKPsi Experience", "Campus Involvement(s)", "Year", "Hometown", "Interests and Hobbies", "Favorite AKPsi Moment"];//different categories from responses

//********************************************************************************************************************** */ 

var spread_sheet = SpreadsheetApp.getActiveSpreadsheet();
var response_sheet = spread_sheet.getSheetByName(response_sheet_name);
var response_vals= response_sheet.getDataRange().getValues()
var i = 0;
var year_arr = [0,0,0,0];
var ebo_arr = [], dir_arr = [], pc_arr = [], active_arr = [],sas_arr = [],loa_arr = [];

for (i = response_vals.length - 1; i > 0; i--)
{
  if (response_vals[i] != null){
  if(response_vals[i][4] == "Eboard")
  {
    if(checkDuplicates(response_vals[i], ebo_arr))
    {
      if (response_vals[i][0] == "President"){
        year_arr.splice(0,1,response_vals[i]);
      }
      else if (response_vals[i][0] == "Executive Vice President"){
        year_arr.splice(1,1,response_vals[i]);
      }
      else if (response_vals[i][0] == "Vice President of Alumni Relations"){
        year_arr.splice(3,1,response_vals[i]);
      }
      else if (response_vals[i][0] == "Vice President of Finance"){
        year_arr.splice(2,1,response_vals[i]);
      }
      else{
        ebo_arr.push(response_vals[i]);
      }
    }
    
  }
  else if(response_vals[i][4] == "Director")
  {
    if(checkDuplicates(response_vals[i], dir_arr)){
      dir_arr.push(response_vals[i])
    }
  }
  else if(response_vals[i][4] == "Pledge Committee")
  {
    if(checkDuplicates(response_vals[i], pc_arr)){
      pc_arr.push(response_vals[i])
    }
  }
  else if(response_vals[i][4] == "SAS"){
    if(checkDuplicates(response_vals[i], sas_arr)){
      sas_arr.push(response_vals[i])
    }
  }
  else if(response_vals[i][4] == "LOA"){
    if(checkDuplicates(response_vals[i], loa_arr)){
    loa_arr.push(response_vals[i])
    }
  }
  else{
    if(checkDuplicates(response_vals[i], active_arr)){
    active_arr.push(response_vals[i]);
    }
  }
}
}

ebo_arr = sortByFirstName(ebo_arr);
dir_arr = sortByFirstName(dir_arr);
pc_arr = sortByFirstName(pc_arr);
actives = sortByFirstName(active_arr);
sas_arr = sortByFirstName(sas_arr);
loa_arr = sortByFirstName(loa_arr);

ebo_arr = year_arr.concat(ebo_arr);
delete response_vals;

var big_list = [ebo_arr,  dir_arr,  pc_arr,  actives,sas_arr,loa_arr];

for (i = 0; i < sheet_name_list.length; i++){
  try{
      var current_sheet = spread_sheet.getSheetByName(sheet_name_list[i]);
      spread_sheet.deleteSheet(current_sheet);
    try{
      spread_sheet.insertSheet().setName(sheet_name_list[i]).getRange("A1:Z500").setBackground("#FFED9E").setFontFamily("Muli");
      current_sheet = spread_sheet.getSheetByName(sheet_name_list[i]);
      
      if (current_sheet != null){
        var row = 0
        var length = big_list[i].length

        for(j = 0; j < length; j++){
          row++;
          row = createEntry(big_list[i][j], row, current_sheet, categories);
          }
        }
    }
    catch{
      Logger.log("Error deleting " + sheet_name_list[i] + " sheet, this may be caused by the sheet not existing");
    }
  }
  catch{
    Logger.log("Error creating " + sheet_name_list[i] + " sheet");
  }

  SpreadsheetApp.getActive().getSheetByName(response_sheet_name).hideSheet();
}
  spread_sheet.getSheetByName(sheet_name_list[0]);
}
function createEntry(current_person,row,sheet, categories){

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
  var link_list = head_link.split("=");
  var id = link_list[1];
  var loop_array = [position, major_etc, work, camp_inv, year, hometown, interests, fav_mom];
  var picture_cell = sheet.getRange(row,1);
  var name_cell = sheet.getRange(row,2);
  
  sheet.setColumnWidths(1,1, 200);
  sheet.setColumnWidths(2,6, 200);
  sheet.getRange(row,2,1,6 ).merge();
  name_cell.setValue(name).setBackground("#6D9EEB").setHorizontalAlignment("center").setFontWeight("bold").setFontSize("16");

  var loop_row = 0;
  var loop_column = 2;

  var mod_val = categories.length/2;
  for(var i = 0; i < categories.length; i++){

    title_cell = sheet.getRange(((loop_row % mod_val) + 1 + row),loop_column);
    title_cell.setValue(categories[i]).setBackground("#c9def8").setFontWeight("bold");
    entry_cell = sheet.getRange(((loop_row % mod_val) + 1+ row),loop_column + 1);
    entry_cell.setValue(loop_array[i]).setBackground("#c9def8");
    sheet.getRange(((loop_row % mod_val + 1) + row),loop_column + 1,1,2 ).merge();

    loop_row ++;

    if(loop_row == mod_val){
      loop_column += 3;
    }
  } 

  var email_cell = sheet.getRange(row+ 5,2);
  sheet.getRange(row + 5,2,1,6 ).merge();
  email_cell.setValue(email).setHorizontalAlignment("center").setBackground("#97BBF1").setFontWeight("bold").setFontSize("10");

  sheet.getRange(row,2,6,6 ).setWrap(true);

/* code to add in headshots, needs lots of fixes
  var a_height = 0;
  for (i = row; i < (row + 6); i++){
    a_height += sheet.getRowHeight(i);
    //Logger.log(sheet.getRowHeight(i))
  }

  sheet.getRange(row,1,6).mergeVertically();//merge column A rows 0-5

  //Logger.log("a_height" + a_height);
  try{
    //Logger.log(sheet.getRowHeight(1));
    var resize_factor = a_height/200; //height/width
    var obj = DriveApp.getFileById(id);
    var image = (obj.imageMediaMetadata && (obj.imageMediaMetadata.width * obj.imageMediaMetadata.height) > 208576 ? obj.thumbnailLink.replace(/=s\d+/, "=s500") : DriveApp.getFileById(id).getBlob());

    var orig_height = ImgApp.getSize(image).height;
    var orig_width = ImgApp.getSize(image).width;
    var orig_factor = orig_height/orig_width;

    if (orig_factor < resize_factor){//width is too large
      var new_width = 200;
      //image = ImgApp.doResize(id, new_width);
    }
    else//height is too large
    {
      //Logger.log("height " + name);
      var new_height = a_height;
      //Logger.log("new_height " + new_height);
      //Logger.log("orig_height " + orig_height);
      var new_width = Math.ceil((orig_width * new_height) / orig_height);
      //Logger.log("new_width " + new_width);
      //Logger.log("orig_width " + orig_width);
      
    }
    image = ImgApp.doResize(id, new_width);
    //Logger.log("width " + (ImgApp.getSize(image.blob).width));
    //Logger.log("height " + (ImgApp.getSize(image.blob).height));
    sheet.insertImage(image.blob, 1, row).setWidth(ImgApp.getSize(image.blob).width).setHeight(ImgApp.getSize(image.blob).height);
  }
 catch{
    Logger.log("Error Inserting Image for " + name);
  }
*/

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

function sortByFirstName(unsorted_list){
  //sort remaining entries by name
  var list_len = unsorted_list.length;
  var sorted_list = [];

  for(i = 0; i < list_len; i++){
    if (sorted_list.length == 0){

      sorted_list.splice(0,0,unsorted_list[0]);
      inserted = true;
    }
    //find alpha first name and place it in the list where it belongs
    else{
      var inserted = false;
      for(j = 0; j < sorted_list.length; j++){
        if(unsorted_list[i][2] < sorted_list[j][2])
        {
          sorted_list.splice(j, 0,unsorted_list[i]);
          inserted = true;
          break;
        }
        
      }
      if (inserted == false){
        sorted_list.push(unsorted_list[i]);
      }
    }

    }
  delete unsorted_list;
  return sorted_list
}
