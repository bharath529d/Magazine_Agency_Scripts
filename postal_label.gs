function generate_postal_shipping_label(data) {
  // Note: The data parameters is passed only when the we generate all the labels together.
  let fileName = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  fileName = fileName.replace(/-/g, '');
  fileName = fileName + "PostalShippingLabel";
  let target_sheet = getSheet("Postal Address Label")
  let required_columns = ["subscription_number", "customer_name", "salutation", "mobile", "phone", "Contact.CF.Area Code", "shipping_address", "shipping_address2", "shipping_city",
  "shipping_state","shipping_zip", "FreeCopies", "No Boxes 1", "Copies Boxes 1", "No Boxes 2", "Copies Boxes 2", "Posting Type", "Despatched through", "Destination Place"]
  let get_column_index = new Map() // indexes of column before preprocessing (we need it to extract only relevant column)
  let final_columns_index = new Map() // column index for the target_sheet after preprocessing (used later in the code).
  if (data) {
    set_data(required_columns, data, target_sheet) // setting the data in the "Postal Address Label" sheet
  } else if (target_sheet.getRange(1, 1).getValue() != "subscription_number") {
    let data_range = getSheet("Subscriptions").getDataRange()
    data = data_range.getValues();
    required_columns.forEach((column_name, index) => {
      get_column_index.set(column_name, data[0].indexOf(column_name))
      final_columns_index.set(column_name, index)
    })
    data = data.filter(row => row[get_column_index.get("Posting Type")].toLowerCase() === "postal");
    data = data.map((row) => {
      let new_row = required_columns.map(column_name => row[get_column_index.get(column_name)]);
      return new_row
    })
    //sorting by subscription_number (ascending)
    data.sort(function (a, b) {
      return a[0].localeCompare(b[0]);
    });

    set_data(required_columns, data, target_sheet) // setting the data in the "Postal Address Label" sheet
  } else {
    console.log("Data Already preprocessed, so let's use it")
    data = target_sheet.getRange(2, 1, target_sheet.getLastRow() - 1, target_sheet.getLastColumn()).getValues()
    console.log("done")
  }

  required_columns.forEach((column_name, index) => {  // getting the index number of each field to extract the column values
    console.log(index)
    final_columns_index.set(column_name, index)
  })
  
  let labels_data = [];
  data.forEach(function (row) {
    let label_data = {};
    let address_data = {};
    address_data.ss_no = get_formatted_ss_no(row[final_columns_index.get("subscription_number")].trim())
    address_data.customer_name = row[final_columns_index.get("customer_name")].trim()
    let salution = row[final_columns_index.get("salutation")]
    if (salution) {
      address_data.salution = salution
    }
    address_data.area_code = row[final_columns_index.get("Contact.CF.Area Code")]
    address_data.shipping_address = row[final_columns_index.get("shipping_address")].trim()
    let shipping_address2 = row[final_columns_index.get("shipping_address2")].trim()
    if (shipping_address2) {
      address_data.shipping_address2 = shipping_address2
    }
    address_data.shipping_city = row[final_columns_index.get("shipping_city")].trim()
    address_data.shipping_state = row[final_columns_index.get("shipping_state")].trim()
    address_data.shipping_zip = row[final_columns_index.get("shipping_zip")]
    label_data.nbox1 = row[final_columns_index.get("No Boxes 1")]
    label_data.ncopies_box1 = row[final_columns_index.get("Copies Boxes 1")]
    let nbox2 = row[final_columns_index.get("No Boxes 2")]
    if (nbox2) {
      label_data.nbox2 = nbox2
    }
    let ncopies_box2 = row[final_columns_index.get("Copies Boxes 2")]
    if (ncopies_box2) {
      label_data.ncopies_box2 = ncopies_box2
    }
    label_data.address_data = address_data
    labels_data.push(label_data);
  })

  let all_labels_data = [];

  function to_number(value) {
    if (typeof value === "string") {
      console.log("Converted to number")
      value = Number(value)
      return value
    }
    return value
  }

  labels_data.forEach((label_data) => {
    let nbox1 = to_number(label_data.nbox1)
    let nbox2 = to_number(label_data.nbox2) || 0
    let nbox = nbox1 + nbox2
    // console.log(`types: ${typeof nbox}, ${typeof nbox1}, ${typeof nbox2}`)
    // console.log(`nbox: ${typeof nbox}, ${typeof nbox1}, ${typeof nbox2}`)
    for (let i = 0; i < nbox1; i++) {
      let label_data_map = new Map();
      label_data_map.address_data = label_data.address_data;
      label_data_map.bundle_no = `${label_data.ncopies_box1}`
      if (nbox > 1) {
        label_data_map.bundle_no += `/${nbox}`
      }
      all_labels_data.push(label_data_map)
    }
    for (let i = 0; i < nbox2; i++) {
      let label_data_map = new Map();
      label_data_map.address_data = label_data.address_data;
      label_data_map.bundle_no = `${label_data.ncopies_box2}/${nbox}`
      all_labels_data.push(label_data_map)
    }
  })
  all_labels_data.forEach((data) => {
    console.log(data)
  })
  let html = getTableHtml(all_labels_data, "postal_label_template.html");
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
