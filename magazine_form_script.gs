function show_form() {
  let html = HtmlService.createHtmlOutputFromFile('magazine_details_form')
    .setTitle('Magazine Details')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function save_form_data(data) {
  let props = PropertiesService.getScriptProperties();
  // Save all fields as key-value
  for (let key in data) {
    props.setProperty(key, data[key]);
  }
  return 'Data saved successfully!';
}

function get_form_data() {
  let props = PropertiesService.getScriptProperties();
  let keys = [
    'magazine_name',
    'magazine_rate',
    'date',
    'weight_per_copy',
    'extra_weight',
    'first_100gm_rate',
    'per_100gm_rate',
    'prepared_by_name'
  ];
  let result = {};
  keys.forEach(function(key) {
    result[key] = props.getProperty(key) || '';
  });
  return result;
}
