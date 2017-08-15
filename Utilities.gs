function showFormKey() {
    var current_time = new Date();
    var last_check = ScriptProperties.getProperty("last_check");
    if (last_check == null) {
        last_check = current_time;
        ScriptProperties.setProperty("last_check", last_check);
    }
    var formUrl = ss.getFormUrl();
    var formKey = formUrl.match(/forms\/d\/([^\/]+)/);
    //var formKey = formUrl.match(/formkey=(.*)/);
    dashboardSheet.getRange('C17').setValue(formKey[1]);
}

function isEmail_(email) {
    var x = email.toString();
    var atpos = x.indexOf("@");
    var dotpos = x.lastIndexOf(".");
    if (atpos < 1 || dotpos < atpos + 2 || dotpos + 2 >= x.length) return false;
}

function scrubData_() {
    var sheet = ss.getSheetByName('Subscribers');
    var data = sheet.getDataRange().getValues();
    var scrubbed_data = new Array();
    // If someone wants to unsubscribe, find his email address and remove it from the mailing list
    for (i in data) {
        if (data[i][2] == 'Yes') {
            for (j in data) {
                if (data[j][1] == data[i][1]) {
                    data[j][1] = 'To delete';
                }
            }
        }
    }
    // Remove duplicates and rows with a wrong email address
    for (i in data) {
        var row = data[i];
        var copy_row = true;
        if (isEmail_(row[1]) == false) {
            copy_row = false;
        }
        if (row[2] == 'Yes') {
            copy_row = false;

        }
        for (j in scrubbed_data) {
            if (row[1] == scrubbed_data[j][1]) {
                copy_row = false;
            }
        }
        if (copy_row) {
            scrubbed_data.push(row);
        }
    }
    var headers = sheet.getRange(1, 1, 1, 3).getValues();
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 3).setValues(headers);
    if (scrubbed_data.length != 0) {
        sheet.getRange(2, 1, scrubbed_data.length, scrubbed_data[0].length).setValues(scrubbed_data);
        dashboardSheet.getRange('G15').setValue(scrubbed_data.length);
    }
    return scrubbed_data;
}
