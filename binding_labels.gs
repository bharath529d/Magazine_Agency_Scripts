function generate_binding_label() {
  // Note: The data parameters is passed only when the we generate all the labels together.
  let fileName = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  fileName = fileName.replace(/-/g, '');
  fileName = fileName + "BindingLabels";
  let required_columns = []
  let get_column_index = new Map() // indexes of column before preprocessing (we need it to extract only relevant column)
  let final_columns_index = new Map() // column index for the target_sheet after preprocessing (used later in the code).
  let data_range =  getSheet("Sheet1").getDataRange()//SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange()
  data = data_range.getValues();
  
  data[0].forEach((column_name, i) => {
    get_column_index.set(column_name,i)
    if(column_name.toLowerCase().includes("area name")){
      final_columns_index.set(column_name, i)
      required_columns.push(column_name)
    }else if (column_name.toLowerCase().includes("binding labels")){
      final_columns_index.set("no of labels", i)
      required_columns.push(column_name)
    }
  })
  
  data = data.filter(row => row[get_column_index.get("Area Name")].toLowerCase() !== "");
  data = data.map((row) => {
    let new_row = required_columns.map(column_name => row[get_column_index.get(column_name)]);
    return new_row
  })

  data.shift()

  let labels_data = [];
  data.forEach(function (row) {
    let label_data = {};
    area_name = row[final_columns_index.get("Area Name")]
    console.log(final_columns_index.get("Area Name"))
    console.log(area_name)
    details = area_name.split(":")
    label_data.area_code = details[0].trim()
    label_data.area_name = details[1].trim()
    if(details[2]){
      label_data.area_name = label_data.area_name + " :"
      label_data.pin_code = details[2].trim()
    }
    label_data.label_count = row[final_columns_index.get("no of labels")]
    labels_data.push(label_data);
  })

  let html = getTableHtml(labels_data, "binding_labels_template.html");
  let folder = DriveApp.getFoldersByName("Address Labels");
  if (folder.hasNext()) {
    folder = folder.next();
  }
  else {
    folder = DriveApp.createFolder("Address Labels");
  }
  let file = folder.createFile(fileName + ".html", html, MimeType.HTML);
  console.log("File created " + file.getId());
}
