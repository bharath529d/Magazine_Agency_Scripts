function generate_all_shipping_labels(){
  if(is_sheet_exists("Subscriptions")){
    let required_columns = ["subscription_number","customer_name","salutation","mobile","phone","Contact.CF.Area Code","shipping_address","shipping_address2","shipping_city","shipping_zip","FreeCopies","No Boxes 1","Copies Boxes 1","No Boxes 2","Copies Boxes 2","Posting Type","Despatched through","Destination Place"]
    let get_column_index = new Map();
    let final_columns_index = new Map();
    let data = getSheet("Subscriptions").getDataRange().getValues();
    required_columns.forEach((column_name, index) => {
      get_column_index.set(column_name, data[0].indexOf(column_name))
      final_columns_index.set(column_name, index)
    })
    
    data.shift()
    data = data.map((row) => {
      let new_row = required_columns.map(column_name => row[get_column_index.get(column_name)]);
      return new_row
    })
    data.sort(function (a, b) {
      return a[0].localeCompare(b[0]);
    });
    let postal_data = []
    let roadways_data = []
    postal_type_index = final_columns_index.get("Posting Type")
    data.forEach(row => {
      if(row[postal_type_index].toLowerCase() === "postal"){
        postal_data.push(row)
      }else if(row[postal_type_index].toLowerCase() === "roadways"){
        roadways_data.push(row)
      }
    }
    )
    generate_postal_shipping_label(postal_data)
    generate_roadways_shipping_label(roadways_data)
    create_door_delivery_labels()
  }else{
    throw new Error("Import the excel data");
  }
}

/*
function generate_all_labels(){
  if(is_sheet_exists("Subscriptions")){
    let required_columns = ["subscription_number","customer_name","salutation","mobile","phone","Contact.CF.Area Code","shipping_address","shipping_address2","shipping_city","shipping_zip","FreeCopies","No Boxes 1","Copies Boxes 1","No Boxes 2","Copies Boxes 2","Posting Type","Despatched through","Destination Place"]
    let get_column_index = new Map();
    let final_columns_index = new Map();
    let data = getSheet("Subscriptions").getDataRange().getValues();
    required_columns.forEach((column_name, index) => {
      get_column_index.set(column_name, data[0].indexOf(column_name))
      final_columns_index.set(column_name, index)
    })
    
    data.shift()
    data = data.map((row) => {
      let new_row = required_columns.map(column_name => row[get_column_index.get(column_name)]);
      return new_row
    })
    data.sort(function (a, b) {
      return a[0].localeCompare(b[0]);
    });
    let postal_data = []
    let roadways_data = []
    postal_type_index = final_columns_index.get("Posting Type")
    data.forEach(row => {
      if(row[postal_type_index].toLowerCase() === "postal"){
        postal_data.push(row)
      }else if(row[postal_type_index].toLowerCase() === "roadways"){
        roadways_data.push(row)
      }
    }
    )
    generate_postal_shipping_label(postal_data)
    generate_roadways_shipping_label(roadways_data)
  }else{
    throw new Error("Import the excel data");
  }
}*/
