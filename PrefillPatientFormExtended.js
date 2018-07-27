
//See - https://nauzadk.wordpress.com/2012/10/02/working-with-commandaction-and-urlaction/

//TODO: Modularise and Wrap all function calls in try / catch error handler.


function addScheduleForExistingPatient(itemId, listId, siteUrl) {
    console.log("addScheduleForExistingPatient - ItemId = " + itemId + ", ListId = " + listId + ", SiteUrl = " + siteUrl);

    var siteServerRelativeUrl = _spPageContextInfo.siteServerRelativeUrl;
    var webServerRelativeUrl = _spPageContextInfo.webServerRelativeUrl;

    console.log(siteServerRelativeUrl);
    console.log(webServerRelativeUrl);

    var myData = initialiseMyData(siteUrl, webServerRelativeUrl, siteServerRelativeUrl);

    // Retrieve Patient Details
    var properties = ['EMR_PatientID', 'EMR_PatientFirstName', 'EMR_PatientLastName'];
    getListItemProperties(listId, itemId, properties).done(function (listItem, properties) {
        readPatientDetails(listItem, properties, myData);
        console.log("Patient ID = " + myData.EMR_PatientID);

        url = webServerRelativeUrl + "/Lists/Schedule/NewForm.aspx?EMR_PatientID=" + myData.EMR_PatientID +
            "&EMR_PatientFirstName=" + encodeURIComponent(myData.EMR_PatientFirstName) +
            "&EMR_PatientLastName=" + encodeURIComponent(myData.EMR_PatientLastName);

        url = url + "&Source=" + window.location.href;

        window.location.href = url;
        //return window.open(webServerRelativeUrl + '/SitePages/Patient%27s%20Forms.aspx?PatientID=' + myData.EMR_PatientID);
    });
}

function showPatientForms(itemId, listId, siteUrl) {
    console.log("showPatientForms - ItemId = " + itemId + ", ListId = " + listId + ", SiteUrl = " + siteUrl);

    var siteServerRelativeUrl = _spPageContextInfo.siteServerRelativeUrl;
    var webServerRelativeUrl = _spPageContextInfo.webServerRelativeUrl;

    console.log(siteServerRelativeUrl);
    console.log(webServerRelativeUrl);

    var myData = initialiseMyData(siteUrl, webServerRelativeUrl, siteServerRelativeUrl);

    // Retrieve Patient Details
    var properties = ['EMR_PatientName', 'EMR_PatientID'];
    getListItemProperties(listId, itemId, properties).done(function (listItem, properties) {
        readPatientDetails(listItem, properties, myData);
        console.log("Patient ID = " + myData.EMR_PatientID);

        window.location.href = webServerRelativeUrl + '/SitePages/Patient%27s%20Forms.aspx?PatientID=' + myData.EMR_PatientID;
        //return window.open(webServerRelativeUrl + '/SitePages/Patient%27s%20Forms.aspx?PatientID=' + myData.EMR_PatientID);
    });
}


// Works from a single select or multiple select.
// Create corresponding items in the FormsToPrint document library.
function generatePatientIntakeForms(itemId, listId, siteUrl) {
    console.log("generatePatientIntakeForms - ItemId = " + itemId + ", ListId = " + listId + ", SiteUrl = " + siteUrl);

    var ctx = SP.ClientContext.get_current();
    var selectedItems = SP.ListOperation.Selection.getSelectedItems(ctx);

    var myItems = '';
    if (selectedItems.length === 0) {
        selectedItems[0] = { id: itemId };
    }

    var siteServerRelativeUrl = _spPageContextInfo.siteServerRelativeUrl;
    var webServerRelativeUrl = _spPageContextInfo.webServerRelativeUrl;

    var myData = initialiseMyData(siteUrl, webServerRelativeUrl, siteServerRelativeUrl);

    if (myData.siteCode.length === 0) {
        alert("Unable to determine susbsite code.");
        return;
    }

    //NOTE: To batch print using an intermediate document libray use addItemToFormsToPrint()

    //Need async 'waterfall loop' to submit print requests for all forms for each patient, one patient at a time.
    //See: http://stackoverflow.com/questions/15504921/asynchronous-loop-of-jquery-deferreds-promises

    //This functionality (concurrent print requests) is available in case the print server is ever scaled out to multiple servers.
    var concurrentPrintRequests = 0;

    //begin the chain by resolving a new $.Deferred
    var dfd = $.Deferred().resolve();

    // use a forEach to create a closure freezing selectedItem
    selectedItems.forEach(function (selectedItem) {

        //Another closure this time freezing myData...
        var myData = initialiseMyData(siteUrl, webServerRelativeUrl, siteServerRelativeUrl);

        // add to the $.Deferred chain with $.then() and re-assign
        dfd = dfd.then(function () {
            return getPatientAndFormDetails(selectedItem.id, listId, myData, siteServerRelativeUrl, webServerRelativeUrl)
				.then(function () {
				    var waitModal;

				    // add to the $.Deferred chain with $.then() and re-assign
				    dfd = dfd.then(function () {
				        if (concurrentPrintRequests) {
				            waitModal = showGeneratingFormsDialog(myData, "");
				            return printPatientIntakeFormsConcurrent(myData);
				        } else {
				            // perform async operation and return its promise
				            return printPatientIntakeForms(myData);
				        }
				    });

				    dfd.done(function () {
				        if (concurrentPrintRequests) {
				            waitModal.close();
				        }
				    });
				});
        });
    });

}

function showGeneratingFormsDialog(myData, friendlyName) {
    var msg = getMsg(myData, friendlyName);
    return SP.UI.ModalDialog.showWaitScreenWithNoClose("Please wait...", msg);;
}

function getMsg(myData, friendlyName) {
    var msg;
    if (friendlyName !== "") {
        msg = "Generating form \"" + friendlyName + "\" for " + myData.EMR_PatientName;
    } else if (myData.forms.length === 1) {
        msg = "Generating form for " + myData.EMR_PatientName;
    } else {
        msg = "Generating " + myData.forms.length + " forms for " + myData.EMR_PatientName;
    }
    return msg;
}

// Issues multiple print requests for a patient concurrently
function printPatientIntakeFormsConcurrent(myData) {
    console.log("in printPatientIntakeFormsConcurrent");

    var deferreds = [];
    for (var i = 0; i < myData.forms.length; i++) {
        var formValue = myData.forms[i];
        var formId = formValue.get_lookupValue();
        deferreds.push(callPrefillPatt(myData, formId));
    }

    return $.when.apply($, deferreds);
}

// Issues multiple print requests for a patient one form at a time (to reduce risk of server timeouts)
function printPatientIntakeForms(myData) {
    console.log("in printPatientIntakeForms");

    //Need async 'waterfall loop' to submit print requests for each form one at a time.
    //See: http://stackoverflow.com/questions/15504921/asynchronous-loop-of-jquery-deferreds-promises
    var waitModal;

    var dfd = $.Deferred().resolve();
    myData.forms.forEach(function (formValue) {

        var formId = formValue.get_lookupValue();

        dfd = dfd.then(function () {

            var friendlyName = findFriendlyName(myData, formId);
            waitModal = showGeneratingFormsDialog(myData, friendlyName);

            return callPrefillPatt(myData, formId);
        });

        dfd.done(function () {
            waitModal.close();
        });

        dfd.fail(function () {
            waitModal.close();
        });

    });

    return dfd;
}


function getPatientAndFormDetails(itemId, listId, myData, siteServerRelativeUrl, webServerRelativeUrl) {
    console.log("getPatientAndFormDetails - ItemId = " + itemId + ", ListId = " + listId);

    //Read New Patient Intake Forms
    var properties = ['SiteCode', 'PracticeID', 'NewPatientIntakeForms'];
    return retrievePracticeDetails(siteServerRelativeUrl, myData.siteCode, properties)
    .then(function (items) {
        return getPracticeDetailsAndPracticeForms(items, myData, true);
    }
	)
	.then(function () {
	    //Get friendly form names.
	    var properties = ['Title', 'DppFormDefn_FriendlyFormId'];
	    return retrieveFriendlyFormNames(siteServerRelativeUrl, myData.siteCode, properties);
	}
	)
	.then(function (items) {
	    return getFriendlyFormNames(items, myData, true);
	}
	)
	.then(function () {
	    // Retrieve Patient Details
	    var properties = ['EMR_PatientName', 'EMR_DOB', 'EMR_PatientID', 'EMR_PatientFirstName', 'EMR_PatientLastName'];
	    return getListItemProperties(listId, itemId, properties).done(function (listItem, properties) {
	        readPatientDetails(listItem, properties, myData);
	    });
	}
	);
}

function initialiseMyData(siteUrl, webServerRelativeUrl, siteServerRelativeUrl) {
    var myData = {};
    myData.isSSL = siteUrl.startsWith("https");
    myData.siteCode = getSiteCode(webServerRelativeUrl, siteServerRelativeUrl);
    return myData;
}

function deleteDocument() {
    var docTitle = "EMR_FinancialPolicy2_Jones_Alex_19671012_1234.txt";
    var clientContext = new SP.ClientContext();
    var oWebsite = clientContext.get_web();
    var fileUrl = _spPageContextInfo.webServerRelativeUrl + "/FormsToPrint/" + docTitle;
    this.fileToDelete = oWebsite.getFileByServerRelativeUrl(fileUrl);
    this.fileToDelete.deleteObject();
    clientContext.executeQueryAsync(
    Function.createDelegate(this, this.onQuerySucceeded),
    Function.createDelegate(this, this.onQueryFailed)
    );
}
function onQuerySucceeded() {
    alert("Document successfully deleted!");
}
function onQueryFailed(sender, args) {
    alert('Request failed. ' + args.get_message() +
    '\n' + args.get_stackTrace());
}

function addItemToFormsToPrint(itemId, listId, myData, siteServerRelativeUrl, webServerRelativeUrl) {
    console.log("addItemToFormsToPrint - ItemId = " + itemId + ", ListId = " + listId);

    //TODO: - Read New Patient Intake Forms
    var properties = ['SiteCode', 'PracticeID', 'Forms'];
    retrievePracticeDetails(siteServerRelativeUrl, myData.siteCode, properties)
    .then(function (items) {
        return getPracticeDetailsAndPracticeForms(items, myData, false);
    }
	)
	.then(function () {
	    // Retrieve Patient Details
	    var properties = ['EMR_PatientName', 'EMR_DOB', 'EMR_PatientID', 'EMR_PatientFirstName', 'EMR_PatientLastName'];
	    return getListItemProperties(listId, itemId, properties);
	}
	)
	.then(function (listItem, properties) {
	    return readPatientDetails(listItem, properties, myData);
	}
	)
	.then(function () {
	    return deleteItemsFromFormsToPrintDocumentLibrary(myData, webServerRelativeUrl);
	}
	)
	.done(function () {
	    addItemsToFormsToPrintDocumentLibrary(myData).then(SP.UI.Notify.addNotification(myData.EMR_PatientName + " - " + myData.forms.length.toString() + " forms queued for printing", false), logError);
	}
	).catch(function (error) {
	    console.log(error);
	});
}

function formatTitle(myData, formName) {
    //FormName_LastName_FirstName_DOB_PatientID.txt

    //Convert DOB to expected format.
    var dob = JSON.parse(JSON.stringify(myData.EMR_DOB));
    var patientDOB = formatDOB(dob);

    return formName +
				"_" + myData.EMR_PatientLastName +
				"_" + myData.EMR_PatientFirstName +
				"_" + patientDOB +
				"_" + myData.EMR_PatientID +
				".txt";
}
function addSingleItemToFormsToPrint(clientContext, oList, myData, formName) {
    var fileName = formatTitle(myData, formName);

    console.log(fileName);

    var fileCreateInfo = new SP.FileCreationInformation();
    fileCreateInfo.set_url(fileName);
    fileCreateInfo.set_content(new SP.Base64EncodedByteArray());
    var fileContent = "Print generation not started.";

    for (var i = 0; i < fileContent.length; i++) {
        fileCreateInfo.get_content().append(fileContent.charCodeAt(i));
    }

    var newFile = oList.get_rootFolder().get_files().add(fileCreateInfo);
    var myListItem = newFile.get_listItemAllFields();
    myListItem.set_item("Params", "These are the params")

    myListItem.update();
}

function addItemsToFormsToPrintDocumentLibrary(myData) {
    var clientContext = SP.ClientContext.get_current();
    var oWebsite = clientContext.get_web();
    var oList = oWebsite.get_lists().getByTitle("FormsToPrint");

    //TODO: Get form names from daily list.
    for (var i = 0; i < myData.forms.length; i++) {
        var formValue = myData.forms[i];
        var formId = formValue.get_lookupValue();

        console.log("formId = " + formId + ", lookupId = " + formValue.get_lookupId());
        addSingleItemToFormsToPrint(clientContext, oList, myData, formId);
    }

    var deferred = $.Deferred();
    clientContext.executeQueryAsync(function () {
        deferred.resolve();
    },
		function (sender, args) {
		    deferred.reject(args.get_message());
		}
	);
    return deferred.promise();
}

function deleteSingleItemFromFormsToPrint(clientContext, myData, formName, oWebsite, webServerRelativeUrl) {
    var fileName = formatTitle(myData, formName);

    console.log(fileName);

    var fileUrl = webServerRelativeUrl + "/FormsToPrint/" + fileName;
    var fileToDelete = oWebsite.getFileByServerRelativeUrl(fileUrl);
    fileToDelete.deleteObject();
    myData.filesToDelete.push(fileToDelete);
}

// Delete any exisiting print requests for the same form and patient.
function deleteItemsFromFormsToPrintDocumentLibrary(myData, webServerRelativeUrl) {
    var clientContext = SP.ClientContext.get_current();
    var oWebsite = clientContext.get_web();

    myData.filesToDelete = [];

    //TODO: Get form names from daily list.
    for (var i = 0; i < myData.forms.length; i++) {
        var formValue = myData.forms[i];
        var formId = formValue.get_lookupValue();

        console.log("formId = " + formId + ", lookupId = " + formValue.get_lookupId());
        deleteSingleItemFromFormsToPrint(clientContext, myData, formId, oWebsite, webServerRelativeUrl);
    }

    var deferred = $.Deferred();
    clientContext.executeQueryAsync(function () {
        deferred.resolve();
    },
		function (sender, args) {
		    deferred.reject(args.get_message());
		}
	);

    return deferred.promise();
}

function prefillScheduleFormExtended(itemId, listId, siteUrl) {
    console.log("prefillScheduleFormExtended - ItemId = " + itemId + ", ListId = " + listId + ", SiteUrl = " + siteUrl);
    prefillFormExtended(itemId, listId, siteUrl, "Schedule");
}
function prefillPatientFormExtended(itemId, listId, siteUrl) {
    console.log("prefillPatientFormExtended - ItemId = " + itemId + ", ListId = " + listId + ", SiteUrl = " + siteUrl);
    prefillFormExtended(itemId, listId, siteUrl, "Patient");
}

function prefillFormExtended(itemId, listId, siteUrl, listType) {
    var siteServerRelativeUrl = _spPageContextInfo.siteServerRelativeUrl;
    var webServerRelativeUrl = _spPageContextInfo.webServerRelativeUrl;

    var myData = {};
    myData.isSSL = siteUrl.startsWith("https");
    myData.siteCode = getSiteCode(webServerRelativeUrl, siteServerRelativeUrl);
    if (myData.siteCode.length === 0) {
        alert("Unable to determine susbsite code.");
        return;
    }

    // If printing from the Schedule list we need to lookup the Patient details via the lookup field.
    var patientListItemID;
    //and also now the Location and Provider details.
    var locationListItemID;
    var providerListItemID;

    var properties = ['SiteCode', 'PracticeID', 'Forms'];
    retrievePracticeDetails(siteServerRelativeUrl, myData.siteCode, properties)
    .then(function (items) {
        return getPracticeDetailsAndPracticeForms(items, myData, false);
    }
	)
	.then(function () {
	    //Get friendly form names.
	    var properties = ['Title', 'DppFormDefn_FriendlyFormId'];
	    return retrieveFriendlyFormNames(siteServerRelativeUrl, myData.siteCode, properties);
	}
	)
	.then(function (items) {
	    return getFriendlyFormNames(items, myData, true);
	}
	)
    .then(function () {
        if (listType === "Schedule") {
            // Need two reads to indirect through lookup field to the Patient details.
            // See: https://sharepoint.stackexchange.com/questions/61663/access-additional-lookup-fields-using-javascript-and-csom
            var properties = ['EMR_Provider_Lookup', 'EMR_Location_Lookup', 'EMR_AppointmentDate', 'Lookup_EMR_PatientID'];
        	return getListItemProperties(listId, itemId, properties);
        }
    }
	)
    .then(function (listItem, properties) {
        if (listType === "Schedule") {
            readScheduleDetails(listItem, properties, myData);
            patientListItemID = readSchedulePatientListItemID(listItem, 'Lookup_EMR_PatientID');
            providerListItemID = readScheduleProviderListItemID(listItem, 'EMR_Provider_Lookup');
            locationListItemID = readScheduleLocationListItemID(listItem, 'EMR_Location_Lookup');
        } else {
            patientListItemID = itemId;
        }
    }
	)
	.then(function () {
	    var properties = ['EMR_PatientName', 'EMR_DOB', 'EMR_PatientID', 'EMR_PatientFirstName', 'EMR_PatientLastName'];
	    return getListItemByListTitleAndItemID("Patients", patientListItemID, properties);
	    }
	)
    .then(function (listItem, properties) {
        return readPatientDetails(listItem, properties, myData);
        }
    )
    .then(function () {
        if (providerListItemID) {
            var properties = ['EMR_Provider'];
            return getListItemByListTitleAndItemID("Providers", providerListItemID, properties);
        }
    }
	)
	.then(function (listItem, properties) {
	    if (providerListItemID) {
	        return readProviderDetails(listItem, properties, myData);
	    }
	}
    )
    .then(function () {
        if (locationListItemID) {
            var properties = ['EMR_Location'];
            return getListItemByListTitleAndItemID("Locations", locationListItemID, properties);
        }
    }
	)
	.then(function (listItem, properties) {
	    if (locationListItemID) {
	        return readLocationDetails(listItem, properties, myData);
	    }
	}
    )
	.done(function () {
	    showFormsToSelect(myData);
	}
	).catch(function (error) {
	    console.log(error);
	});

}

function getSiteCode(webServerRelativeUrl, siteServerRelativeUrl) {
    return webServerRelativeUrl.substr(siteServerRelativeUrl.length + 1);
}

function retrievePracticeDetails(siteServerRelativeUrl, siteCode, propertiesToInclude) {
    console.log("calling retrievePracticeDetails");

    var ctx = new SP.ClientContext(siteServerRelativeUrl);
    var web = ctx.get_web();
    var list = web.get_lists().getByTitle('Practices');

    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><Query><Where><Eq><FieldRef Name=\'SiteCode\'/>' +
        '<Value Type=\'Text\'>' + siteCode + '</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>');

    var items = list.getItems(camlQuery);
    var includeExpr = 'Include(' + propertiesToInclude.join(',') + ')';

    ctx.load(items, includeExpr);

    var d = $.Deferred();
    ctx.executeQueryAsync(function () {
        var result = items.get_data().map(function (i) {
            return i.get_fieldValues();
        });
        d.resolve(result);
    },
    function (sender, args) {
        d.reject(args);
    });

    return d.promise();
}

function retrieveFriendlyFormNames(siteServerRelativeUrl, siteCode, propertiesToInclude) {
    console.log("calling retrieveFriendlyFormNames");

    var ctx = new SP.ClientContext(siteServerRelativeUrl);
    var web = ctx.get_web();
    var list = web.get_lists().getByTitle('DppFormDefnLibrary');

    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View></View>');

    var items = list.getItems(camlQuery);
    var includeExpr = 'Include(' + propertiesToInclude.join(',') + ')';

    ctx.load(items, includeExpr);

    var d = $.Deferred();
    ctx.executeQueryAsync(function () {
        var result = items.get_data().map(function (i) {
            return i.get_fieldValues();
        });
        d.resolve(result);
    },
    function (sender, args) {
        d.reject(args);
    });

    return d.promise();
}

function getPracticeDetailsAndPracticeForms(items, myData, isNewPatientIntakeForms) {
    console.log("calling getPracticeDetailsAndPracticeForms");
    //console.log(JSON.stringify(items));
    if (items.length != 1) {
        alert("Found " + items.length + " practices matching SiteCode " + myData.siteCode);
        return;
    }

    if (items[0].PracticeID == null) {
        alert("Unable to find matching PracticeID for SiteCode " + myData.siteCode);
        return;
    }

    myData.practiceID = items[0].PracticeID;

    if (isNewPatientIntakeForms) {
        if (items[0].NewPatientIntakeForms.length === 0) {
            alert("Unable to find matching new patient intake forms for SiteCode " + myData.siteCode);
            return;
        } else {
            myData.forms = items[0].NewPatientIntakeForms;
        }
    } else {
        if (items[0].Forms.length === 0) {
            alert("Unable to find matching forms for SiteCode " + myData.siteCode);
            return;
        } else {
            myData.forms = items[0].Forms;
        }
    }

    console.log(JSON.stringify(myData.forms));
    console.log("found " + myData.practiceID + " matching SiteCode " + myData.siteCode);
}

function getFriendlyFormNames(items, myData) {
    console.log("in getFriendlyFormNames");

    myData.dppFormDefn = [];

    for (var i = 0; i < items.length; i++) {
        console.log(items[i].Title); //FormID
        console.log(items[i].DppFormDefn_FriendlyFormId);
        if (items[i].DppFormDefn_FriendlyFormId.toUpperCase === "UNKNOWN") {
            //Use FormID as friendly name
            myData.dppFormDefn[items[i].Title] = items[i].Title;
        } else {
            myData.dppFormDefn[items[i].Title] = items[i].DppFormDefn_FriendlyFormId;
        }
    }
}

function findFriendlyName(myData, formId) {
    if (myData.dppFormDefn[formId] === undefined) {
        //Not found - shouldn't happen (unless form has been deleted from DppFormDef library.
        console.log(formId + " not found in DppFormDefn library");
        return formId;
    }

    return (myData.dppFormDefn[formId]);
}

/* Show a modalDialog with the contents of divPrintSelectedFormsContainer */
function showFormsToSelect(myData) {
    // showModalDialog removes the element passed in from the DOM
    // so we save a copy in a closure to add back later
    var reAddClonedFormWrapper = (function () {
        var privatecopyOfPrintSelectedForms = $("#divPrintSelectedFormsContainer").clone(true, true);
        var execute = function () {
            jQuery('#s4-bodyContainer').append(privatecopyOfPrintSelectedForms);
        };

        return {
            execute: execute
        }
    })();

    var output = [];

    /*
        var selectValues = {
            "test 1": "test 1",
            "test 2": "test 2",
            "test 3": "test 3",
            "test 4": "test 4",
            "test 5": "test 1",
            "test 6": "test 2",
            "test 7": "test 3",
            "test 8": "test 4"
        };
    
        $.each(selectValues, function (key, value) {
            output.push('<option value="' + key + '">' + value + '</option>');
        });
    */

    for (var i = 0; i < myData.forms.length; i++) {
        var formValue = myData.forms[i];
        var formId = formValue.get_lookupValue();

        var friendlyName = findFriendlyName(myData, formId);
        output.push('<option value="' + formId + '">' + friendlyName + '</option>');

        console.log("formId = " + formId);
    }

    $('#selectForms').html(output.join(''));

    var selectHeight = parseInt($("#selectForms option").length + 1) * 30;
    var dialogHeight = selectHeight + 40;

    divPrintSelectedFormsContainer.style.display = "inline";

    var options = {
        html: divPrintSelectedFormsContainer,
        title: 'Print Patient Forms',
        width: 380,
        height: dialogHeight,
        args: JSON.stringify(myData),
        autoSize: false,
        dialogReturnValueCallback: reAddClonedFormWrapper.execute
    };

    modalDialog = SP.UI.ModalDialog.showModalDialog(options);

    // See http://stackoverflow.com/questions/12640828/setting-a-multiselect-chosen-boxs-height to improve behaviour of height.
    var mySelectForms = $('#selectForms').chosen({
        placeholder_text_multiple: "Select forms...",
    });

    /*
	$('#selectForms').on('change', function(evt, params) {
		_resizeModalDialog(evt, params);
	});
	*/

}

//This version left in in case we ever scale out to multiple print servers.
function printSelectedFormsConcurrent(data) {

    var myData = JSON.parse(data);
    //console.log(myData);

    var selectedFormsArray = $("#selectForms").val() || [];
    console.log("selected forms: " + selectedFormsArray.join());

    SP.UI.ModalDialog.commonModalDialogClose(SP.UI.DialogResult.OK);

    //For each selected form make ajax call to print server. TODO: Needs testing!
    //TODO: async behaviour may give better UX. See: e.g. https://developers.google.com/web/fundamentals/getting-started/primers/promises, https://elgervanboxtel.nl/site/blog/xmlhttprequest-extended-with-promises

    var msg;
    if (selectedFormsArray.length === 1) {
        msg = "Generating form for " + myData.EMR_PatientName;
    } else {
        msg = "Generating " + selectedFormsArray.length + " forms for " + myData.EMR_PatientName;
    }

    var waitModal = SP.UI.ModalDialog.showWaitScreenWithNoClose("Please wait...", msg);

    var deferreds = [];
    for (var i = 0; i < selectedFormsArray.length; i++) {
        deferreds.push(callPrefillPatt(myData, selectedFormsArray[i]));
    }

    $.when.apply($, deferreds).done(function () {
        waitModal.close();
    }
	);
}

function printSelectedForms(data) {

    var myData = JSON.parse(data);
    //console.log(myData);

    var selectedFormsArray = $("#selectForms").val() || [];
    console.log("selected forms: " + selectedFormsArray.join());

    SP.UI.ModalDialog.commonModalDialogClose(SP.UI.DialogResult.OK);

    //For each selected form make ajax call to print server. 
    //Async behaviour gives better UX. See: e.g. 
    //https://developers.google.com/web/fundamentals/getting-started/primers/promises
    //https://elgervanboxtel.nl/site/blog/xmlhttprequest-extended-with-promises

    //Need async 'waterfall loop' to submit print requests for each form one at a time.
    //See: http://stackoverflow.com/questions/15504921/asynchronous-loop-of-jquery-deferreds-promises
    var waitModal;
    var formCount = 0;

    var dfd = $.Deferred().resolve();
    selectedFormsArray.forEach(function (formId) {

        var thisFormCount = ++formCount;
        var msg;
        if (selectedFormsArray.length === 1) {
            msg = "Generating form for " + myData.EMR_PatientName;
        } else {
            msg = myData.EMR_PatientName.trim() + ": Generating form " + thisFormCount + " out of " + selectedFormsArray.length;
        }

        dfd = dfd.then(function () {
            waitModal = SP.UI.ModalDialog.showWaitScreenWithNoClose("Please wait...", msg);
            return callPrefillPatt(myData, formId);
        });

        dfd.done(function () {
            waitModal.close();
        });

        dfd.fail(function () {
            waitModal.close();
        });

    });
}

function callPrefillPatt(myData, formID) {

    console.log("in callPrefillPatt: formID = " + formID);

    var params = "PracticeID=" + myData.practiceID + "&WRSPatientID=" + myData.EMR_PatientID + "&formid=" + formID + "&PatientName=" + myData.EMR_PatientName;

    //Use friendly form name to generate downloaded PDF name
    //If there is a requirement to download the PDF using the friendly form name then look it up here!


    //CRQ-101749: Add new fields to support new naming of PDF image files for export.
    //CRQ-101749: Field names registered via the EMR_Template form and must be added to ALL new EMR digital forms.

    //Convert DOB to expected format.
    var dob = JSON.parse(JSON.stringify(myData.EMR_DOB));
    var patientDOB = formatDOB(dob);

    params = params + "&PatientDOB=" + patientDOB + "&PatientFirstName=" + myData.EMR_PatientFirstName + "&PatientLastName=" + myData.EMR_PatientLastName
    params = params + "&EMRLocation=" + (myData.EMR_Location || "") + "&EMRProvider=" + (myData.EMR_Provider || "");

    params = appendFormSpecificParams(params, myData, formID);
    params = params + appendDownloadFileName(params, myData, formID, patientDOB);
    params = params.replace(/ /g, "+");
    console.log(params);

    return doCallPrefillPatt(myData.isSSL, params);
}

function appendDownloadFileName(params, myData, formID, patientDOB) {
    var downloadFileName = formID + "_" + myData.EMR_PatientLastName + "_" + myData.EMR_PatientFirstName + "_" + patientDOB;
    return params + "&pdfDownloadName=" + downloadFileName;
}

function appendFormSpecificParams(params, myData, formID) {
    var updatedParams = params;

    if (formID.startsWith("ENT_Specialty_Care_Survey")) {
        updatedParams = updatedParams + "&FirstNameInitial=" + myData.EMR_PatientFirstName + " " + myData.EMR_PatientLastName.charAt(0);
        updatedParams = updatedParams + "&TodaysDate=" + todaysDate(); //TODO: Use Appointment time?
        updatedParams = updatedParams + "&DoctorName=" + (myData.EMR_Provider || "");
    }

    if (formID.startsWith("EMR_PatientExperienceSurveyTest")) {
        updatedParams = updatedParams + "&PracticeName=" + (myData.EMR_Location || "") + "&DoctorName=" + (myData.EMR_Provider || "");
    }

    return updatedParams;
}

function formatDOB(dob) {
    // Format is 1953-06-05T04:00:00.000Z, return YYYYMMDD
    dob = dob.replace(/-/g, '');
    dob = dob.substr(0, 8);
    return dob;
}

function todaysDate() {
    var today = new Date();
    var dd = today.getDate();
    var mm = today.getMonth() + 1; //January is 0!
    var yyyy = today.getFullYear();

    if (dd < 10) {
        dd = '0' + dd
    }

    if (mm < 10) {
        mm = '0' + mm
    }

    today = mm + '/' + dd + '/' + yyyy;
    return today;
}

function doCallPrefillPatt(isSSL, params) {
    var url = "http://SITE/URL.php";
    if (isSSL) {
        url = url.replace("http", "https");
    }

    console.log("in doCallPrefillPatt: " + params);

    var deferred = $.Deferred();

    //Based on http://stackoverflow.com/questions/16086162/handle-file-download-from-ajax-post/23797348#23797348
    var xhr = new XMLHttpRequest();
    xhr.open('POST', url, true);
    xhr.responseType = 'arraybuffer';

    xhr.onload = function () {
        if (this.readyState === 4 && this.status === 200) {
            var filename = "";
            var disposition = xhr.getResponseHeader('Content-Disposition');
            if (disposition && disposition.indexOf('attachment') !== -1) {
                var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                var matches = filenameRegex.exec(disposition);
                if (matches != null && matches[1]) filename = matches[1].replace(/['"]/g, '');
            }
            var type = xhr.getResponseHeader('Content-Type');

            var blob = new Blob([this.response], { type: type });
            if (typeof window.navigator.msSaveBlob !== 'undefined') {
                // IE workaround for "HTML7007: One or more blob URLs were revoked by closing the blob for which they were created. These URLs will no longer resolve as the data backing the URL has been freed."
                window.navigator.msSaveBlob(blob, filename);
            } else {
                var URL = window.URL || window.webkitURL;
                var downloadUrl = URL.createObjectURL(blob);

                if (filename) {
                    // use HTML5 a[download] attribute to specify filename
                    var a = document.createElement("a");
                    // safari doesn't support this yet
                    if (typeof a.download === 'undefined') {
                        window.location = downloadUrl;
                    } else {
                        a.href = downloadUrl;
                        a.download = filename;
                        document.body.appendChild(a);
                        a.click();
                    }
                } else {
                    window.location = downloadUrl;
                }

                setTimeout(function () { URL.revokeObjectURL(downloadUrl); }, 1000); // cleanup
            }

            deferred.resolve();
        }
    };

    xhr.setRequestHeader('Content-type', 'application/x-www-form-urlencoded');
    xhr.send(params);
    return deferred;
}

function readScheduleDetails(listItem, properties, myData) {
    console.log("readScheduleDetails:");
    readDataFromProperties(listItem, properties, myData)
    var dateTime = myData['EMR_AppointmentDate'];
    
    //See: https://developer.mozilla.org/en/docs/Web/JavaScript/Reference/Global_Objects/Date/toLocaleString
    // SharePoint date time will be stored in UTC. SharePoint and JSOM map the UTC time to the browser timezone
    // so the date we use here in prefilling will be in the local timezone (which is correct) but will need
    // handling (somehow) on form upload back to SharePoint (since do we know the local timezone of the pen?) 
    // Maybe the form should record some timezone offset from server time?
    // See e.g. http://sadomovalex.blogspot.co.uk/2016/12/get-current-server-date-time-via.html
    //TODO: This should really use the SharePoint web local.
    myData['EMR_AppointmentDate'] = dateTime.toLocaleString('en-US', { hour12: true });
    console.log("SharePoint appt time: " + dateTime.toString());
    console.log("Prefilling appt time: " + myData['EMR_AppointmentDate']);
}

function readPatientDetails(listItem, properties, myData) {
    console.log("readPatientDetails:");
    readDataFromProperties(listItem, properties, myData)
}

function readProviderDetails(listItem, properties, myData) {
    console.log("readProviderDetails:");
    readDataFromProperties(listItem, properties, myData)
}

function readLocationDetails(listItem, properties, myData) {
    console.log("readLocationDetails:");
    readDataFromProperties(listItem, properties, myData)
}

function readDataFromProperties(listItem, properties, myData) {
    for (var i = 0; i < properties.length; i++) {
        myData[properties[i]] = listItem.get_item(properties[i]);
        console.log(properties[i] + " = " + myData[properties[i]]);
    }
}

function readSchedulePatientListItemID(listItem, lookupFieldName) {
    var patientListItemID = getLookupListItem(listItem, lookupFieldName);
    console.log("readSchedulePatientListItemID:" + " = " + patientListItemID);
    return patientListItemID;
}

function readScheduleProviderListItemID(listItem, lookupFieldName) {
    var providerListItemID = getLookupListItem(listItem, lookupFieldName);
    console.log("readScheduleProviderListItemID:" + " = " + providerListItemID);
    return providerListItemID;
}

function readScheduleLocationListItemID(listItem, lookupFieldName) {
    var locationListItemID = getLookupListItem(listItem, lookupFieldName);
    console.log("readScheduleLocationListItemID:" + " = " + locationListItemID);
    return locationListItemID;
}

function getLookupListItem(listItem, lookupFieldName) {
    var lookup = listItem.get_item(lookupFieldName);
    var lookupListItemID = lookup.get_lookupId();
    console.log("readSchedulePatientListItemID:" + " = " + lookupListItemID);
    return lookupListItemID;
}

//NOT TESTED (uses private MS functions)
function _resizeModalDialog() {
    console.log("in _resizeModalDialog");

    // get the top-most dialog
    var dlg = SP.UI.ModalDialog.get_childDialog();

    if (dlg != null) {
        // dlg.$Q_0 - is dialog maximized
        // dlg.get_$b_0() - is dialog a modal

        if (!dlg.$Q_0 && dlg.get_$b_0()) {
            // resize the dialog
            dlg.autoSize();

            var xPos, yPos, //x & y co-ordinates to move modal to...
                win = SP.UI.Dialog.get_$1(), // the very bottom browser window object
                xScroll = SP.UI.Dialog.$1x(win), // browser x-scroll pos
                yScroll = SP.UI.Dialog.$20(win); // browser y-scroll pos

            //SP.UI.Dialog.$1P(win) - get browser viewport width         
            //SP.UI.Dialog.$1O(win) - get browser viewport height
            //dlg.$3_0 - modal's DOM element

            // calculate x-pos based on viewport and dialog width
            xPos = ((SP.UI.Dialog.$1P(win) - dlg.$3_0.offsetWidth) / 2) + xScroll;

            // if x-pos is out of view (content too wide), re-position to left edge + 10px
            if (xPos < xScroll + 10) { xPos = xScroll + 10; }

            // calculate y-pos based on viewport and dialog height
            yPos = ((SP.UI.Dialog.$1O(win) - dlg.$3_0.offsetHeight) / 2) + yScroll;

            // if x-pos is out of view (content too high), re-position to top edge + 10px
            if (yPos < yScroll + 10) { yPos = yScroll + 10; }

            // store dialog's new x-y co-ordinates
            dlg.$K_0 = xPos;
            dlg.$W_0 = yPos;

            // move dialog to x-y pos
            dlg.$p_0(dlg.$K_0, dlg.$W_0);

            // set dialog title bar text width
            dlg.$1b_0();

            // size down the dialog width/height if it's larger than browser viewport
            dlg.$27_0();
        }
    }
}

function initPrefillPrintExtended() {
    // Create a hidden div which will be used by the SP modal dialog
    // TODO: This is added to every SharePoint page. Control this by using the Calculated Column 'hack'?
    // TODO: Add cancel button. Need to position correctly.
    var divPrintSelectedFormsContainer = '<div id="divPrintSelectedFormsContainer" style="display: none;">' +
											'<div id="divDropDownForms" style="float: left; ">' +
												'<select id="selectForms" multiple="multiple" ></select>' +
											'</div>' +
											'<div id="divPrintButton" style="float: right; ">' +
												'<input type="button" value="Print" onclick="printSelectedForms(SP.UI.ModalDialog.get_childDialog().get_args())" />' +
											'</div>' +
										'</div>' +
										'<script></script>';


    jQuery('#s4-bodyContainer').append(divPrintSelectedFormsContainer);
}

// Ensure the SP JSOM is loaded before we worry about our stuff
jQuery(document).ready(function () {
    ExecuteOrDelayUntilScriptLoaded(initPrefillPrintExtended, "sp.js");
});




