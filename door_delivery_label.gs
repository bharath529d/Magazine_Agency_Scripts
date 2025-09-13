var Decimal = DecimalJsLib.Decimal // this is decimal object which is got from decimaljsLib

function create_door_delivery_labels(){
  let target_sheet = getSheet("Door Delivery Label")
  // Columns needed from the sheet to prepare the invoice
  let required_columns = ["subscription_number", "customer_name", "shipping_address", "shipping_address2","shipping_city",
  "shipping_zip", "phone","mobile", "No Boxes 1", "Copies Boxes 1", "No Boxes 2", "Copies Boxes 2", "Free Copies", "Posting Type"] 
  // get_column_index stores indexes(from the Subscription sheet) of the required columns .
  let get_column_index = new Map() // indexes of column before selection of columns from the subscription sheet (we need this to extract only the relevant column)
  // final_columns_index stores the indexes of the required columns after we have only the relevant columns
  let final_columns_index = new Map()
  let data_range = getSheet("Subscriptions").getDataRange()
  // data now stores all the data as a 2d array from the Subscription sheet
  data = data_range.getValues();
  // stores indexes for the columns as planned
  required_columns.forEach((column_name, index) => {
    get_column_index.set(column_name, data[0].indexOf(column_name))
    final_columns_index.set(column_name, index)
  })
  // Filtering only the rows that is needed for preparing the invoice (in our case, we need rows that have either 'postal' or 'more copies' in the Posting type column) 
  data = data.filter(row => {
    let posting_type = (row[get_column_index.get("Posting Type")] || "").toLowerCase();
    return posting_type === "door delivery" 
  });
  // now get extract only the relevannt columns from the rows and stores it in data 
  data = data.map((row) => {
    let new_row = required_columns.map(column_name => row[get_column_index.get(column_name)]);
    return new_row
  })
  //sorting by subscription_number (ascending)
  data.sort(function (a, b) {
      return a[0].localeCompare(b[0]);
  });
  set_data(required_columns, data, target_sheet)
  let labels_data = [];
  data.forEach(function (row) {
    // label_data stores the data need to create the label (we have 3 label per page)
    let label_data = {}; 
    label_data.ss_no = row[final_columns_index.get("subscription_number")].trim()
    label_data.customer_name = row[final_columns_index.get("customer_name")].trim()
    label_data.shipping_address = row[final_columns_index.get("shipping_address")].trim()
    label_data.shipping_address2 = row[final_columns_index.get("shipping_address2")].trim()
    label_data.shipping_city = row[final_columns_index.get("shipping_city")].trim()
    label_data.shipping_zip = row[final_columns_index.get("shipping_zip")].trim()
    label_data.phone = row[final_columns_index.get("phone")]
    label_data.mobile = row[final_columns_index.get("mobile")]
    label_data.free_copies = toInt(row[final_columns_index.get("Free Copies")])
    label_data.nbox1 = toInt(row[final_columns_index.get("No Boxes 1")])
    label_data.ncopies_box1 = toInt(row[final_columns_index.get("Copies Boxes 1")])
    let nbox2 = toInt(row[final_columns_index.get("No Boxes 2")])
    if (nbox2) {
      label_data.nbox2 = nbox2
    }
    let ncopies_box2 = toInt(row[final_columns_index.get("Copies Boxes 2")])
    if (ncopies_box2) {
      label_data.ncopies_box2 = ncopies_box2
    }
    labels_data.push(label_data);
  })

  let htmlTemplate = HtmlService.createTemplateFromFile("door_delivery_template.html");
  htmlTemplate.labels_data = labels_data;
  let ps = PropertiesService.getScriptProperties()
  htmlTemplate.magazine_name = ps.getProperty("magazine_name")
  htmlTemplate.date = ps.getProperty('date')
  htmlTemplate.prepared_by_name = ps.getProperty('prepared_by_name')
  let html = htmlTemplate.evaluate().getContent();
  let folder = DriveApp.getFoldersByName("Address Labels");
  if (folder.hasNext()) {
    folder = folder.next();
  }
  else {
    folder = DriveApp.createFolder("Address Labels");
  }
  let file = folder.createFile("DD_label_" + ps.getProperty('date') + ".html", html, MimeType.HTML);
  console.log("File created " + file.getId());

}

