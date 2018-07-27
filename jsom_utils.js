function getListItemProperties(listId, itemId, properties) {
    var ctx = SP.ClientContext.get_current();
    var list = ctx.get_web().get_lists().getById(listId);
    var listItem = list.getItemById(itemId);
    var deferred = $.Deferred();

    ctx.load(listItem, properties);

    ctx.executeQueryAsync(function () {
        deferred.resolve(listItem, properties);
    },
    function (sender, args) {
        deferred.reject(args.get_message());
    });
    return deferred.promise();
}

//TODO: Only include fields which are actually needed. Map properties to include argument in ctx.load() call.
function getListItemsCamlQuery(listId, camlQuery, properties, usefulData) {
    var ctx = SP.ClientContext.get_current();
    var list = ctx.get_web().get_lists().getById(listId);
    var items = list.getItems(camlQuery);
    var deferred = $.Deferred();

    ctx.load(items);
    ctx.executeQueryAsync(function () {
        deferred.resolve(items, properties, usefulData);
    },
    function (sender, args) {
        deferred.reject(args.get_message(), usefulData);
    });
    return deferred.promise();
}

function getListItemsCamlQueryInclude(listId, camlQuery, properties, includeFields, usefulData) {
    var ctx = SP.ClientContext.get_current();
    var list = ctx.get_web().get_lists().getById(listId);
    var items = list.getItems(camlQuery);
    var deferred = $.Deferred();

    ctx.load(items, includeFields);
    ctx.executeQueryAsync(function () {
        deferred.resolve(items, properties, usefulData);
    },
    function (sender, args) {
        deferred.reject(args.get_message(), usefulData);
    });
    return deferred.promise();
}

function getListItemsByListTitleCamlQueryInclude(listTitle, camlQuery, properties, includeFields, usefulData) {
    var ctx = SP.ClientContext.get_current();
    var list = ctx.get_web().get_lists().getByTitle(listTitle);
    var items = list.getItems(camlQuery);
    var deferred = $.Deferred();

    ctx.load(items, includeFields);
    ctx.executeQueryAsync(function () {
        deferred.resolve(items, properties, usefulData);
    },
    function (sender, args) {
        deferred.reject(args.get_message(), usefulData);
    });
    return deferred.promise();
}

function getListItemByListTitleAndItemID(listTitle, itemId, properties) {
    var deferred = $.Deferred();
    var ctx = SP.ClientContext.get_current();
    var list = ctx.get_web().get_lists().getByTitle(listTitle);
    var listItem = list.getItemById(itemId);
    var deferred = $.Deferred();

    ctx.load(listItem, properties);

    ctx.executeQueryAsync(function () {
        deferred.resolve(listItem, properties);
    },
    function (sender, args) {
        deferred.reject(args.get_message());
    });
    return deferred.promise();
}

function getListItems(listTitle, includeFields, usefulData) {
    var deferred = $.Deferred();
    var ctx = SP.ClientContext.get_current();
    var list = ctx.get_web().get_lists().getByTitle(listTitle);
    var items = list.getItems(SP.CamlQuery.createAllItemsQuery());
    ctx.load(items, includeFields);
    ctx.executeQueryAsync(function () {
        deferred.resolve(items, usefulData);
    },
    function (sender, args) {
        deferred.reject(args.get_message(), usefulData);
    });
    return deferred.promise();
}

function updateListItems(items, itemProperties) {
    var deferred = $.Deferred();
    var ctx = items.get_context();

    //prepare to update list items 
    var itemCount = 0;
    items.get_data().forEach(function (item) {
        itemCount++;
        for (var propName in itemProperties) {
            item.set_item(propName, itemProperties[propName])
        }
        item.update();
    });

    //submit request
    ctx.executeQueryAsync(function () {
        deferred.resolve(itemCount + ' item(s) updated'); //TODO: items not needed here ("updated" with a count would be better)
    },
    function (sender, args) {
        deferred.reject(args.get_message());
    });
    return deferred.promise();
}

function logError(error) {
    console.log('An error occured: ' + error);
}

function logErrorWithPageRefresh(error) {
    console.log('An error occured: ' + error);

    RefreshPage(SP.UI.DialogResult.OK);
}

function logSuccess(sender, args) {
    console.log('Updated.');
}

function logSuccessWithPageRefresh(message) {
    if (typeof message !== 'undefined') {
        console.log(message);
    }

    RefreshPage(SP.UI.DialogResult.OK);
}

//See: https://www.eliostruyf.com/ajax-refresh-item-rows-in-sharepoint-2013-view/	
function logSuccessWithItemRefresh(sender, args) {
    console.log('Updated.');

    // Set Ajax refresh context
    var evtAjax = {
        currentCtx: ctx,
        csrAjaxRefresh: true
    };

    // If set to false all list items will refresh
    ctx.skipNextAnimation = true;
    // Initiate Ajax Refresh on the list
    AJAXRefreshView(evtAjax, SP.UI.DialogResult.OK);
}

function logErrorWithItemRefresh(error) {
    console.log(error);

    // Set Ajax refresh context
    var evtAjax = {
        currentCtx: ctx,
        csrAjaxRefresh: true
    };

    // If set to false all list items will refresh
    ctx.skipNextAnimation = true;
    // Initiate Ajax Refresh on the list
    AJAXRefreshView(evtAjax, SP.UI.DialogResult.OK);
}

//See: https://stackoverflow.com/questions/901115/how-can-i-get-query-string-values-in-javascript
function getParameterByName(name, url) {
    if (!url) url = window.location.href;
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
        results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
}

//See: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Math/random
function getRandomInt(min, max) {
    min = Math.ceil(min);
    max = Math.floor(max);
    return Math.floor(Math.random() * (max - min)) + min; //The maximum is exclusive and the minimum is inclusive
}