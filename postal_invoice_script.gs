var Decimal = DecimalJsLib.Decimal // this is decimal object which is got from decimaljsLib
function create_invoice(type){
  let target_sheet, file_name,template_name;
  // Columns needed from the sheet to prepare the invoice
  let required_columns = ["subscription_number", "customer_name", "No Boxes 1", "Copies Boxes 1", "No Boxes 2", "Copies Boxes 2", "Posting Type"] 
  // get_column_index stores indexes(from the Subscription sheet) of the required columns .
  let get_column_index = new Map() // indexes of column before selection of columns from the subscription sheet (we need it to extract only relevant column)
  // final_columns_index stores the indexes of the required columns after get have only the relevant columns
  let final_columns_index = new Map()
  let data_range = getSheet("Subscriptions").getDataRange()
  // data now stores all the data as a 2d array from the Subscription sheet
  data = data_range.getValues();
  // stores indexes for the columns as planned
  required_columns.forEach((column_name, index) => {
    get_column_index.set(column_name, data[0].indexOf(column_name))
    final_columns_index.set(column_name, index)
  })

  // Filtering only the rows what is needed for preparing the invoice.
  if(type == 'P'){
    target_sheet = getSheet("Postal Invoice")
    template_name = 'postal_invoice_template.html'
    file_name = 'PostalInvoice'
    data = data.filter(row => {
      let posting_type = (row[get_column_index.get("Posting Type")] || "").toLowerCase();
      return posting_type === "postal" 
    });
  }else if(type == 'M'){
    target_sheet = getSheet("More Copies Invoice")
    template_name = 'more_copies_invoice_template.html'
    file_name = 'MoreCopiesInvoice'
    data = data.filter(row => {
      let posting_type = (row[get_column_index.get("Posting Type")] || "").toLowerCase();
      return posting_type === "more copies";
    });
  }else if(type == 'B'){
    target_sheet = getSheet("Both P & MC Invoices")
    template_name = 'postal_invoice_template.html'
    file_name = 'Both_Postal_More_Copies_Invoices'
    data = data.filter(row => {
      let posting_type = (row[get_column_index.get("Posting Type")] || "").toLowerCase();
      return posting_type === "postal" || posting_type === "more copies";
    });
  }

  
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
  
  //cpy_bdle contains copies_per_bundle as key and no_of_bundles as value
  cpy_bdle = new Map();
  
  data.forEach((row) => {
    let copies_per_bundle = toInt(row[final_columns_index.get("Copies Boxes 1")])
    let no_of_bundles = toInt(row[final_columns_index.get("No Boxes 1")])
    let prev_no_of_bdle;
    if(copies_per_bundle > 0 && no_of_bundles > 0){
      prev_no_of_bdle = toInt(cpy_bdle.get(copies_per_bundle))
      cpy_bdle.set(copies_per_bundle, prev_no_of_bdle + no_of_bundles)
    }
    copies_per_bundle = toInt(row[final_columns_index.get("Copies Boxes 2")])
    no_of_bundles = toInt(row[final_columns_index.get("No Boxes 2")])
    if(copies_per_bundle > 0 && no_of_bundles > 0){
      prev_no_of_bdle = toInt(cpy_bdle.get(copies_per_bundle))
      cpy_bdle.set(copies_per_bundle, prev_no_of_bdle + no_of_bundles)
    }
  })

  // grouped_bundles contains array of a map where each map contains copies_per_bundle, weight per bundle, postage due, no. of bundles, total postage due, total copies

  let grouped_bundles = []
  let grand_total_details = new Map();
  let grand_total_no_of_bundle = 0
  let grand_total_postage_due= new Decimal(0)
  let grand_total_copies = 0

  let ps = PropertiesService.getScriptProperties()
  for (const [copies_per_bundle, no_of_bundles] of cpy_bdle) {
    let bundle = new Map();
    bundle.set("copies_per_bundle", copies_per_bundle)
    bundle.set("no_of_bundles", no_of_bundles)
    let weight_per_copy = ps.getProperty("weight_per_copy")
    let extra_weight = ps.getProperty("extra_weight")
    // Multiplying 1000 to convert the kg into gm for easy calculation
    let weight_per_bundle = new Decimal(weight_per_copy).times(copies_per_bundle).plus(extra_weight).toDecimalPlaces(3, Decimal.ROUND_HALF_EVEN)
    bundle.set("weight_per_bundle", weight_per_bundle.times(1000).toString())
    let first_100gm_rate = ps.getProperty("first_100gm_rate")
    let per_100gm_rate = ps.getProperty("per_100gm_rate")
    let postage_due = calculate_postage_due(weight_per_bundle, new Decimal(first_100gm_rate), new Decimal(per_100gm_rate))
    bundle.set("postage_due", postage_due.toString())
    let total_postage_due = postage_due.times(no_of_bundles)
    bundle.set("total_postage_due", total_postage_due.toString())
    let total_copies = no_of_bundles * copies_per_bundle
    bundle.set("total_copies", total_copies)
    grouped_bundles.push(bundle)

    // grand total details
    grand_total_no_of_bundle += no_of_bundles
    grand_total_postage_due = grand_total_postage_due.plus(total_postage_due)
    grand_total_copies += total_copies
  }
  grand_total_details.set('grand_total_no_of_bundle',grand_total_no_of_bundle)
  grand_total_details.set('grand_total_postage_due',grand_total_postage_due.toString())
  grand_total_details.set('grand_total_copies',grand_total_copies)

  grouped_bundles.sort((bundle1, bundle2) => bundle1.get("copies_per_bundle") - bundle2.get("copies_per_bundle"))

  let htmlTemplate = HtmlService.createTemplateFromFile(template_name);
  htmlTemplate.grouped_bundles = grouped_bundles;
  htmlTemplate.grand_total_details = grand_total_details;
  htmlTemplate.magazine_name = ps.getProperty("magazine_name")
  htmlTemplate.date = ps.getProperty('date')
  let html = htmlTemplate.evaluate().getContent();
  let folder = DriveApp.getFoldersByName("All Invoices");
  if (folder.hasNext()) {
    folder = folder.next();
  }
  else {
    folder = DriveApp.createFolder("All Invoices");
  }
  let file = folder.createFile(file_name + ps.getProperty('date') + ".html", html, MimeType.HTML);
  console.log("File created " + file.getId());
}

/**
 * @param {Decimal} weight_per_bundle  // in kilograms
 * @param {Decimal} first_100gm_rate
 * @param {Decimal} per_100gm_rate     // per additional 100g or part thereof
 * @returns {Decimal}
 */
function calculate_postage_due(weight_per_bundle, first_100gm_rate, per_100gm_rate) {
  // Convert kg to grams
  const grams = weight_per_bundle.times(1000); // Decimal

  // If weight is zero or negative, you can decide policy; here we return 0
  if (grams.lte(0)) {
    return new Decimal(0);
  }

  // First 100g covered by base rate
  if (grams.lte(100)) {
    return new Decimal(first_100gm_rate);
  }

  // Compute extra grams beyond first 100g
  const extra = grams.minus(100); // Decimal

  // Each started 100g costs per_100gm_rate
  const extraUnits = extra.div(100).ceil(); // Decimal
  let postage_due =  new Decimal(first_100gm_rate).plus(per_100gm_rate.times(extraUnits));

  return postage_due
}

function create_more_copies_invoice(){
 create_invoice('M')
}

function create_postal_invoice(){
 create_invoice('P')
}

function create_both_invoices(){
 create_invoice('B')
}






