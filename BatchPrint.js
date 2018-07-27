//BatchPrint.js
//TODO: Add into single global EMR object.
var EMR_LockManager = (function () {
    var locks = {};
    var masterLockId = null;

    return {
        lockCount: function () {
            var name;
            var lockCount = 0;
            for (name in locks) {
                if (locks.hasOwnProperty(name)) {
                    if (locks[name] === true) {
                        lockCount++;
                    }
                }
            }

            return lockCount;
        },

        getLock: function (lockId, masterLock) {
            if (masterLockId !== null) {
                return false;
            }

            var id = lockId.toString();
            console.log("lockId.toString() = " + id);
            if (typeof locks[id] === 'undefined' || locks[id] === false) {
                locks[id] = true;
                if (masterLock) {
                    masterLockId = id;
                }
                return true;
            }

            return false;
        },

        //Return lock only if no other locks are currently held.
        getMasterLock: function (lockId) {
            if (EMR_LockManager.lockCount() == 0) {
                return EMR_LockManager.getLock(lockId, true);
            } else {
                return false;
            }
        },

        releaseLock: function (lockId) {
            var id = lockId.toString();
            locks[id] = false;
            if (id === masterLockId) {
                masterLockId = null;
                console.log("Released master lock");
            }
        },

        //NOTE: We don't treat initialise all as a true lock, since this function runs every time the page loads.
        isLocked: function () {
            return EMR_LockManager.lockCount() > 0 && masterLockId !== "initialiseAll";
        },
    };
}());

function getScheduleLock(apptid) {
    return EMR_LockManager.getLock(apptid, false);
}

function getMasterLock(lockId) {
    return EMR_LockManager.getMasterLock(lockId);
}

function releaseLock(lockId) {
    return EMR_LockManager.releaseLock(lockId);
}

function releaseMasterLock(lockId) {
    return EMR_LockManager.releaseLock(lockId);
}

function updatesInProgress() {
    return EMR_LockManager.isLocked();
}

//This will run on every page in SharePoint - NOT now since called from Calculated Column 
//(MagicLink column) =DppFormBase_FormStatus&"<img src=""/_layouts/images/blank.gif"" onload=""{"&"initBatchPrint();}"">"
//Initialise onClick events for the different Print Set Status Values
function initBatchPrint() {
    //Flag to ensure only called once per page load.
    if (typeof initBatchPrint.counter === 'undefined') {
        // It has not... perform the initialization
        initBatchPrint.counter = 0;

        console.log("in initBatchPrint: updating hrefs - " + initBatchPrint.counter++);

        $("img[id^='SelectSchedule_']").click(function () {
            //togglePrint($(this).data("apptid"), $(this).data("set"));
            togglePrint($(this).attr('id'));
            return true;
        });

        $("img[id^='DeselectSchedule_']").click(function () {
            togglePrint($(this).attr('id'));
            return true;
        });

        $("img[id^='DownloadPDF']").click(function () {
            downloadPDF($(this).attr('id'));
            return true;
        });

        $("img[id^='ErrorSchedule_']").click(function () {
            clearError($(this).attr('id'));
            return true;
        });

        // We want to detect if the user refreshes or navigates away from the page.
        $(window).on('beforeunload', function () {
            return confirmPageReload();
        });

        ExecuteOrDelayUntilScriptLoaded(initialiseAllEmptyStatusFields, "sp.js");
    }

    //$("a[href$='.txt']").css("background-color", "yellow");
}

function confirmPageReload() {
    if (updatesInProgress()) {
        return "Changes have been made that have not yet been saved back to the sserver. If you leave now some of those changes may be lost.";
    }
}

// Note: Works only on schedules actually visible in the current view.
//       AND only FormSets that are visible in current view.
//       We do this so that using a view filter like Today's schedules gives the required behaviour.
//       But also this means that any view filter together with Select all will work in the same way.
//       Downside is: more than 99 schedules on one day, will need to page through and Select All on each page.
//       This is different to the Download Today's behaviour which matches all schedules on date.
function selectAll() {
    var usefulData = {};
    usefulData.lockId = "toggleAll";

    //Need to get a master lock i.e. make sure no other JSOM calls are still updating items.
    if (!getMasterLock(usefulData.lockId)) {
        alert("Waiting for server to update schedules. Please try again in a few moments.");
        console.log("Cannot get lock - ignored toggle all event");
        return;
    }

    var found = false;
    var selectedApptIDs = [];
    $("img[data-action='Not Set']").each(function () {
        setImgSelected($(this));
        selectedApptIDs.push($(this).data("apptid"));
        found = true;
        return;
    });

    //Note: The UI may not have been up to date with the server. 
    //      We can no longer make a JSOM call to update all items,
    //      since Select all now works on current displayed schedules only.
    //      Instead let's just do a page refresh.
    if (!found) {
        SP.UI.Notify.addNotification("No items displayed in current view to select.", false);
        releaseLock(usefulData.lockId);
        return;
    }

    //We need to find ALL the form sets on this page (not just Queued or Not Set). This is because
    //subsequent JSOM calls want to find schedules that may not be visible in this page of the current view.
    var formSets = getAllFormSets();

    console.log(formSets);
    console.log(selectedApptIDs);
    toggleAllStatus(selectedApptIDs, formSets, "Not Set", "Queued", usefulData);
}

function deselectAll() {
    var usefulData = {};
    usefulData.lockId = "toggleAll";

    //Need to get a master lock i.e. make sure no other JSOM calls are still updating items.
    if (!getMasterLock(usefulData.lockId)) {
        alert("Waiting for server to update schedules. Please try again in a few moments.");
        console.log("Cannot get lock - ignored toggle all event");
        return;
    }

    var found = false;
    var deselectedApptIDs = [];
    $("img[data-action='Queued']").each(function () {
        setImgNotSelected($(this));
        deselectedApptIDs.push($(this).data("apptid"));
        found = true;
        return;
    });

    //Note: The UI may not have been up to date with the server. 
    //      We can no longer make a JSOM call to update all items,
    //      since Select all now works on current displayed schedules only.
    if (!found) {
        SP.UI.Notify.addNotification("No items displayed in current view to deselect.", false);
        releaseLock(usefulData.lockId);
        return;
    }

    //We need to find ALL the form sets on THIS page (not just Queued or Not Set). This is because
    //subsequent JSOM calls want to find schedules that may not be visible in this page of the current view.
    var formSets = getAllFormSets();

    console.log(formSets);
    console.log(deselectedApptIDs);
    toggleAllStatus(deselectedApptIDs, formSets, "Queued", "Not Set", usefulData);
}


//We have to update all blank EMR_Print_Set[A-Z]_Status choice values to "Not Set"
//This is unfortunately necessary since adding new items with the EMR_Schedule content type
//does not populate the the EMR_BathPrint ct default fields.
function initialiseAllEmptyStatusFields() {
    var usefulData = {};
    usefulData.lockId = "initialiseAll";

    //Need to get a master lock i.e. make sure no other JSOM calls are still updating items.
    if (!getMasterLock(usefulData.lockId)) {
        console.log("Cannot get lock - ignored toggle all event");
        return;
    }

    var found = false;
    var deselectedApptIDs = [];
    $("img[data-action='Not Set']").each(function () {
        setImgNotSelected($(this));
        deselectedApptIDs.push($(this).data("apptid"));
        found = true;
        return;
    });

    //Note: The UI may not have been up to date with the server. 
    //      We can no longer make a JSOM call to update all items,
    //      since Select all now works on current displayed schedules only.
    if (!found) {
        releaseLock(usefulData.lockId);
        return;
    }

    //We need to find ALL the form sets on THIS page (not just Queued or Not Set). This is because
    //subsequent JSOM calls want to find schedules that may not be visible in this page of the current view.
    var formSets = getAllFormSets();

    console.log(formSets);
    console.log(deselectedApptIDs);
    toggleAllStatus(deselectedApptIDs, formSets, null, "Not Set", usefulData);
}



function getAllFormSets() {
    var formSets = Object.create(null);
    $("img[data-set^='Set']").each(function () {
        formSets[$(this).data("set")] = true;
        return;
    });

    return formSets;
}

function downloadTodays() {
    var dateString = "<Today />";
    downloadByAppointmentDate(dateString);
}

function downloadTomorrows() {
    var today = new Date();
    var tomorrowDateTime = new Date();
    tomorrowDateTime.setDate(today.getDate() + 1);
    var dateString = tomorrowDateTime.toISOString();
    console.log("Note: tomorrow's date may possibly be incorrect depending on client timezone... " + dateString);
    downloadByAppointmentDate(dateString);
}

function downloadByAppointmentDate(dateString) {

    var queryString = buildDateBasedCamlQueryString(dateString);

    if (queryString === "") {
        SP.UI.Notify.addNotification("No PDFs available to download.", false);
        return;
    }

    console.log(queryString);

    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(queryString);

    var fieldProperties = {};
    var usefulData = {};
    var includeFields = "Include(EMR_AppointmentID)";
    var listId = _spPageContextInfo.pageListId;

    var waitModal = SP.UI.ModalDialog.showWaitScreenWithNoClose("Please wait...", "");

    //May need to display wait dialog... (Particularly if server is slow e.g. first request after iis restart)
    getListItemsCamlQueryInclude(listId, camlQuery, fieldProperties, includeFields, usefulData)
    .then(doDownloadSelected, logError)
    .then(function () { waitModal.close(); }, function () { waitModal.close(); });
}

// No longer used, since need to check if thre is anything available to download first.
function OLD_downloadTodays() {
    // An empty set of apptIDs means ALL
    //TODO: We may need to work out the number of form pages per ready formset and then just request the 
    //TODO: appropriate number of schedules' PDFs, so as not to overload the server - not easy to do - or the client printing.
    var apptIDs = [];
    var apptIDsParam = JSON.stringify(apptIDs);

    downloadAll('Today', apptIDsParam);
}

function OLD_downloadTomorrows() {

    // An empty set of apptIDs means ALL
    //TODO: We may need to work out the number of form pages per ready formset and then just request the 
    //TODO: appropriate number of schedules' PDFs, so as not to overload the server - not easy to do - or the client printing.
    var apptIDs = [];
    var apptIDsParam = JSON.stringify(apptIDs);

    downloadAll('Tomorrow', apptIDsParam);
}



function downloadSelected() {

    var ctx = SP.ClientContext.get_current();
    var selectedItems = SP.ListOperation.Selection.getSelectedItems(ctx);

    //Need to use CSOM to get selected items appointment IDs.
    var queryString = buildSelectedItemsCamlQueryString(selectedItems);

    console.log(queryString);

    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(queryString);

    var fieldProperties = {};
    var usefulData = {};
    var includeFields = "Include(EMR_AppointmentID)";
    var listId = _spPageContextInfo.pageListId;

    var waitModal = SP.UI.ModalDialog.showWaitScreenWithNoClose("Please wait...", "");

    //May need to display wait dialog... (Particularly if server is slow e.g. first request after iis restart)
    getListItemsCamlQueryInclude(listId, camlQuery, fieldProperties, includeFields, usefulData)
    .then(doDownloadSelected, logError)
    .then(function () { waitModal.close(); }, function () { waitModal.close(); });
}

function doDownloadSelected(items, itemProperties, usefulData) {
    var apptIDs = [];
    var itemCount = 0;
    items.get_data().forEach(function (item) {
        itemCount++;
        apptIDs.push(item.get_item('EMR_AppointmentID'));
    });

    console.log(itemCount);

    if (itemCount === 0) {
        SP.UI.Notify.addNotification("No PDFs available to download.", false);
        return;
    }

    console.log(JSON.stringify(apptIDs));
    var apptIDsParam = JSON.stringify(apptIDs);

    downloadAll('', apptIDsParam);
}

function buildDateBasedCamlQueryString(dateString) {
    var formSets = getAllFormSets();
    var fieldNames = buildStatusFieldNames(formSets);
    if (fieldNames.length === 0) {
        return "";
    }

    var formSetQueryString = buildStatusQueryString(fieldNames, "Ready", "", true);

    //NOTE: The query string format created by CAML Builder doesn't work i.e. [Today] so instead use <Today />
    var dateQueryString = '<Eq><FieldRef Name="EMR_AppointmentDate" /><Value IncludeTimeValue="FALSE" Type="DateTime">' + dateString + '</Value></Eq>';

    return "<View><Query><Where><And>" + dateQueryString + formSetQueryString + "</And></Where></View></Query>";
    //return "<View><Query><Where>" + dateQueryString + "</Where></View></Query>";
}

function buildSelectedItemsCamlQueryString(selectedItems) {
    var formSets = getAllFormSets();
    var fieldNames = buildStatusFieldNames(formSets);

    var formSetQueryString = buildStatusQueryString(fieldNames, "Ready", "", true);

    var idQueryString = buildIdQueryString(selectedItems);

    return "<View><Query><Where><And>" + idQueryString + formSetQueryString + "</And></Where></View></Query>";
}

function buildIdQueryString(selectedItems) {
    var queryString = "<In>" +
                        "<FieldRef Name='ID' />" +
                        "<Values>";

    var i;
    for (i = 0; i < selectedItems.length; i++) {
        queryString += "<Value Type='Counter'>";
        queryString += selectedItems[i].id;
        queryString += "</Value>";
    }

    return queryString + "</Values>" +
                        "</In>";
}

// Note: Works on all schedules in the list, not just what's visible in current view.
//       BUT only FormSets that are visible in current view.  
//       Unlike Select / Deselect All which only affect visible items.
//
//       param 'day' is "Today" or "Tomorrow" or "" 
//
function downloadAll(day, apptIDsParam) {

    var formSets = getAllFormSets();
    var formSetNames = buildFormSetNames(formSets);

    var formSetsParam = JSON.stringify(formSetNames);
    console.log(formSetsParam);

    // 'http://internal.css.local/emr/Goshen/_layouts/15/EMR/DownloadSchedules.aspx?printJobID=ALL&formSets={'SetA','SetB'}&day=Today&listID={b7929b55-d560-4c83-830e-1a77a62eb96b}';
    var url = _spPageContextInfo.webAbsoluteUrl + '/' + _spPageContextInfo.layoutsUrl +
                        '/EMR/DownloadSchedules.aspx?apptIDs=' + apptIDsParam + '&formSets=' + formSetsParam + '&day=' + day + '&listID=' + _spPageContextInfo.pageListId;

    console.log(url);
    window.open(url, '_blank');
}

function downloadPDF(imgID) {
    var img = $('#' + imgID);

    //Make query string params consistent with multiple selection case.
    var formSetNames = [];
    formSetNames.push(img.data("set"));
    var formSetsParam = JSON.stringify(formSetNames);

    var apptIDs = [];
    apptIDs.push(img.data("apptid"));
    var apptIDsParam = JSON.stringify(apptIDs);

    // 'http://internal.css.local/emr/goshen/_layouts/15/EMR/DownloadSchedules.aspx?printJobID=123&formSet=EMR_Print_SetA_Action&listID=GUID'
    var url = _spPageContextInfo.webAbsoluteUrl + '/' + _spPageContextInfo.layoutsUrl +
        '/EMR/DownloadSchedules.aspx?apptIDs=' + apptIDsParam + '&formSets=' + formSetsParam + '&listID=' + _spPageContextInfo.pageListId;

    console.log(url);
    window.open(url, '_blank');
}

function clearError(imgID) {
    //TODO: Display nice dialog with error message, OK, Cancel buttons...
    //Relies on data-action for Error status being set to Queued. (It is now set to error).
    //togglePrint(imgID);
}

//Load the correct image without a slow page refresh.
function setImgNotSelected(img) {
    img.attr("data-action", "Not Set");
    img.attr("src", "/_layouts/15/images/MeticulusPrintFormGrey.png");
}

//Load the correct image without a slow page refresh.
function setImgSelected(img) {
    img.attr("data-action", "Queued");
    img.attr("src", "/_layouts/15/images/MeticulusPrintFormGreen.png");
}


// Update all status fields (e.g. EMR_Print_SetA_Status, EMR_Print_SetB_Status etc.).
// Formsets object contains name of any form sets which exist (in the view, since found using jQuery)  
// and have at least one schedule that needs updating.
function toggleAllStatus(apptIDs, formSets, oldStatus, newStatus, usefulData) {
    var listId = _spPageContextInfo.pageListId;

    var fieldProperties = getFieldProperties(formSets, newStatus);
    var fieldNames = buildStatusFieldNames(formSets);

    console.log(fieldNames);

    var apptIDsQueryString = buildApptIdQueryString(apptIDs);
    console.log(apptIDsQueryString);

    var fieldStatusQueryString = buildStatusQueryString(fieldNames, oldStatus, '', true);
    console.log(fieldStatusQueryString);

    var fullQueryString = "<View><Query><Where><And>" + apptIDsQueryString + fieldStatusQueryString + "</And></Where></View></Query>";
    console.log(fullQueryString);

    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(fullQueryString);

    getListItemsCamlQuery(listId, camlQuery, fieldProperties, usefulData)
    .then(processSchedule)
    .then(handleSuccess, handleFailure);
}

function buildApptIdQueryString(apptIDs) {
    var queryString = "<In>" +
                        "<FieldRef Name='EMR_AppointmentID' />" +
                        "<Values>";

    var i;
    for (i = 0; i < apptIDs.length; i++) {
        queryString += "<Value Type='Text'>";
        queryString += apptIDs[i];
        queryString += "</Value>";
    }

    return queryString + "</Values>" +
                        "</In>";
}

//Note: Called recursively
function buildStatusQueryString(fieldNames, status, queryString, isSubClauseOnly) {
    while (fieldNames.length > 0) {
        var fieldName = fieldNames.pop();
        var fieldQuery = fieldStatusEqual(fieldName, status);
        if (queryString === '') {
            queryString = fieldQuery;
        } else {
            queryString = '<Or>' + queryString + fieldQuery + '</Or>';
        }
        return buildStatusQueryString(fieldNames, status, queryString, isSubClauseOnly);
    }
    if (isSubClauseOnly) {
        return queryString;
    } else {
        return '<View><Query><Where>' + queryString + '</Where></Query></View>';
    }
}

function fieldStatusEqual(fieldName, status) {
    if (status === null) {
        return '<IsNull><FieldRef Name="' + fieldName + '" /></IsNull>';
    } else {
        return '<Eq><FieldRef Name="' + fieldName + '" /><Value Type="Choice">' + status + '</Value></Eq>';
    }
}



function buildFormSetNames(formSets) {
    var name;
    var names = [];
    for (name in formSets) {
        names.push(name);
    }
    return names;
}

function buildStatusFieldNames(formSets) {
    var name;
    var fieldName;
    var fieldNames = [];
    for (name in formSets) {
        fieldName = "EMR_Print_" + name + "_Status";
        fieldNames.push(fieldName);
    }
    return fieldNames;
}

function getFieldProperties(formSets, newStatus) {
    var fieldProperties = {};
    var name;
    var fieldName;
    for (name in formSets) {
        console.log(name);
        fieldName = "EMR_Print_" + name + "_Status";
        fieldProperties[fieldName] = newStatus;
    }

    return fieldProperties;
}

function togglePrint(imgID) {
    console.log(imgID);
    var img = $('#' + imgID);
    var apptid = img.data("apptid");

    if (!getScheduleLock(apptid)) {
        alert("Waiting for server to update schedule. Please try again in a few moments.");
        console.log("Cannot get lock: togglePrint already running - ignored event");
        return;
    }

    console.log("togglePrint - let's do it");

    var usefulData = {};
    usefulData.imgID = imgID;
    usefulData.lockId = apptid;

    // NOTE: Don't use the jQuery img.data("action") method here, since this does not get updated until after a page refresh.
    if (img.attr("data-action") === "Queued") {
        setImgNotSelected(img);
        usefulData.toggleOn = false;
        doTogglePrint(imgID, "Queued", "Not Set", usefulData);
    } else {
        setImgSelected(img);
        usefulData.toggleOn = true;
        doTogglePrint(imgID, "Not Set", "Queued", usefulData);
    }
}


function doTogglePrint(imgID, currentStatus, newStatus, usefulData) {
    var img = $('#' + imgID);
    var apptid = img.data("apptid"); //Note: data() method OK here since these attrs do not change.
    var formSet = img.data("set");

    console.log("in doTogglePrint: apptid - " + apptid + ", formSet - " + formSet + ', new status = ' + newStatus);
    var listId = _spPageContextInfo.pageListId;

    var fieldName = "EMR_Print_" + formSet + "_Status";
    var fieldProperties = {};
    fieldProperties[fieldName] = newStatus;
    //var fieldProperties = { 'EMR_Print_SetA_Status': 'Not Set' };

    //find matching item in SP list and set it's status to "Not Set"		
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><Query><Where>' +
            '<And>' +
                '<Eq><FieldRef Name="EMR_AppointmentID" /><Value Type="Text">' + apptid + '</Value></Eq>' +
                '<Eq><FieldRef Name="' + fieldName + '" /><Value Type="Choice">' + currentStatus + '</Value></Eq>' +
            '</And>' +
        '</Where></Query></View>');

    getListItemsCamlQuery(listId, camlQuery, fieldProperties, usefulData)
    .then(processSchedule)
    .then(handleSuccess, handleFailure);
}

function handleSuccess(usefulData) {
    releaseLock(usefulData.lockId);
    console.log("success");
}

function handleFailure(error, usefulData) {
    console.log("in handleFailure");
    releaseLock(usefulData.lockId); //If we got here we must have acquired the lock in this browser client.

    //If we have hit a JSOM failure, best approach is to refresh client page, since it appears the data 
    //we are showing in the client is out of sync with the server. This situation occurs with multiple clients.
    /* Not needed now since we do a page refresh,
    var img = $('#' + usefulData.imgID);
    if (usefulData.toggleOn) {
        setImgNotSelected(img);
    } else {
        setImgSelected(img);
    }
    */

    //alert("Update may have failed: please try again.");
    logErrorWithPageRefresh(error);
}

function processSchedule(items, itemProperties, usefulData) {

    var deferred = $.Deferred();

    var ctx = items.get_context();

    var itemCount = 0;
    //prepare to update list items 
    items.get_data().forEach(function (item) {
        itemCount++;
        for (var propName in itemProperties) {
            item.set_item(propName, itemProperties[propName])
        }
        item.update();
    });

    console.log(itemCount);

    //If our query has failed to update any items then this is only an error in the single item case.
    //In the single item case, we always make a JSOM call to update the schedule on the server (even if no change to state).
    //But in the De/Select All case, we only query for specific states and updates those items so it's
    //possible there will be no matches - e.g. if our browser client was out of date.
    if (itemCount == 0 && (usefulData.lockId !== "toggleAll" && usefulData.lockId !== "initialiseAll")) {
        deferred.reject("No matching apptId (" + usefulData.lockId +
                                ") and status found. UI may be out of sync with server. Page refresh required.", usefulData);
    } else {
        //submit request
        ctx.executeQueryAsync(function () {
            deferred.resolve(usefulData);
        },
		function (sender, args) {
		    deferred.reject(args.get_message(), usefulData);
		});
    }

    return deferred.promise();
}

// Ensure the SP JSOM is loaded before we worry about our stuff
jQuery(document).ready(function () {
    //ExecuteOrDelayUntilScriptLoaded(initialiseAllEmptyStatusFields, "sp.js");
});


