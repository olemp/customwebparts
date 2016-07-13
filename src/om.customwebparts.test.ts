namespace OM.CustomWebParts.Test {
    export function SearchTable(webpart: Model.WebPart) {
        var properties = webpart.properties[0];

        var searchQuery = properties["Query"];

        if (properties["OnlyListItems"] === "true") {
            searchQuery += " contentclass:STS_ListItem";
        }
        if (properties["OnlyDocuments"] === "true") {
            searchQuery += " IsDocument:True";
        }

        var managedProperties = properties["Properties"];
        var rowLimit = properties["RowLimit"] || 10;

        jQuery.ajax({
            url: String['format']("{0}/_api/search/query?querytext='{1}'&selectproperties='{2}'&rowlimit={3}", _spPageContextInfo.webAbsoluteUrl, searchQuery, managedProperties, rowLimit),
            type: "get",
            headers: {
                "Accept": "application/json;odata=nometadata"
            },
            success: function (d) {
                jQuery("head").append(String['format']("<link rel='stylesheet' type='text/css' href='{0}' />", "//cdn.datatables.net/1.10.7/css/jquery.dataTables.min.css"));

                var results = d.PrimaryQueryResult.RelevantResults.Table.Rows;
                var managedPropertiesArray = managedProperties.split(",");

                var items = [];

                results.forEach(function (val, id) {
                    var item = {};
                    for (var i = 0; i < managedPropertiesArray.length; i++) {
                        var mp = managedPropertiesArray[i];

                        item[mp] = jQuery.grep(val.Cells, function(cell) {
                            return cell["Key"] === mp;
                        })[0]["Value"];
                    }
                    items.push(item);
                });

                var header = ["<th>", managedPropertiesArray.join("</th><th>"), "</th>"].join("");
                var body = "";
                items.forEach(function (itm) {
                    var itemValues = "";
                    Object.keys(itm).forEach(function(key) {
                        itemValues += ["<td>", itm[key], "</td>"].join("");
                    });

                    body += [
                        "<tr>",
                            itemValues,
                        "</tr>"
                    ].join("");
                });


                webpart.instance.html(["<table id='search-table' style='display:none;'>",          
                    "<thead>",
                    "<tr>",
                    header,
                    "</tr>",
                    "</thead>",
                    "<tbody>",
                    body,
                    "</tbody>",    
                "</table>"].join(""));

                jQuery.getScript('//cdn.datatables.net/1.10.7/js/jquery.dataTables.min.js', function() {
                    $('#search-table')['DataTable']();
                    $('#search-table').show();
                });
            },
            error: function(sender, args) {
                console.log(args);
            }
        });
    }
    export function SearchTableExtended(webpart: Model.WebPart) {
        var properties = webpart.properties[0];

        var searchQuery = properties["Query"];

        if (properties["OnlyListItems"] === "true") {
            searchQuery += " contentclass:STS_ListItem";
        }
        if (properties["OnlyDocuments"] === "true") {
            searchQuery += " IsDocument:True";
        }

        var managedProperties = properties["Properties"];
        var rowLimit = properties["RowLimit"] || 10;

        Util.search(searchQuery, managedProperties, rowLimit, 0, null).done(function (results) {
            jQuery("head").append(String['format']("<link rel='stylesheet' type='text/css' href='{0}' />", "//cdn.datatables.net/1.10.7/css/jquery.dataTables.min.css"));

            var managedPropertiesArray = managedProperties.split(",");
            console.log(managedPropertiesArray);
            var items = [];

            results.forEach(function (val) {
                var item = {};
                for (var i = 0; i < managedPropertiesArray.length; i++) {
                    var mp = managedPropertiesArray[i];

                    item[mp] = jQuery.grep(val.Cells, function (cell) {
                        return cell["Key"] === mp;
                    })[0]["Value"];
                }
                items.push(item);
            });

            var header = ["<th>", managedPropertiesArray.join("</th><th>"), "</th>"].join("");
            var body = "";
            items.forEach(function (itm) {
                var itemValues = "";
                Object.keys(itm).forEach(function (key) {
                    itemValues += ["<td>", itm[key], "</td>"].join("");
                });

                body += [
                    "<tr>",
                    itemValues,
                    "</tr>"
                ].join("");
            });


            webpart.instance.html(["<table id='search-table-extended' style='display:none;'>",
                "<thead>",
                "<tr>",
                header,
                "</tr>",
                "</thead>",
                "<tbody>",
                body,
                "</tbody>",
                "</table>"].join(""));

            jQuery.getScript('//cdn.datatables.net/1.10.7/js/jquery.dataTables.min.js', function () {
                $('#search-table-extended')['DataTable']();
                $('#search-table-extended').show();
            });
        });
    }
    export function YammerEmbedAction(webpart: Model.WebPart) {
        var properties = webpart.properties[0];

        if (properties["Network"] && ["Action"]) {
            webpart.instance.html(String['format']("<div id='{0}' /></div>", 'embedded-feed'));
    

            var embedOptions = {
                container: "#embedded-feed",
                network: properties["Network"],
                action: properties["Action"]
            };

            jQuery.getScript('https://assets.yammer.com/assets/platform_embed.js', function () {
                yam.connect.actionButton(embedOptions);
            });
        } else {
            console.log("You need to specify Network Action for the yammer embed webpart.")
        }
    }
    export function YammerEmbed(webpart: Model.WebPart) {
        var properties = webpart.properties[0];

        if (properties["Network"] && ["FeedType"] && ["FeedId"]) {
            var width = properties["Width"] || "400px";
            var height = properties["Height"] || "800px";
            var promptText = properties["PromptText"] || "Say something..";
            var header = properties["Header"] ? properties["Header"] == "true" : false;
            var footer = properties["Footer"] ? properties["Footer"] == "true" : false;

            webpart.instance.html(String['format']("<div id='{0}' style='width:{1};height:{2}' /></div>", 'embedded-feed', width, height));
    
            var embedOptions = {
                container: "#embedded-feed",
                network: properties["Network"],
                feedType: properties["FeedType"],
                feedId: properties["FeedId"],
                config: {
                    header: header,
                    footer: footer,
                    promptText: promptText
                }
            };

            if (properties["OverrideObjectProperties"] == "true") {
                embedOptions["objectProperties"] = {
                    type: "page",
                    title: _spPageContextInfo.webTitle,
                    description: "",
                    image: "https://mug0.assets-yammer.com/mugshot/images/128x128/FvM5Sp1j7bXl-N9HvKKLjqZ44BCFPxGL"
                }
            }

            jQuery.getScript('https://assets.yammer.com/assets/platform_embed.js', function () {
                yam.connect.embedFeed(embedOptions);
            });
        } else {
            console.log("You need to specify Network, Feed Type and Feed ID for the yammer embed webpart.")
        }
    }
    export function Subwebs(webpart: Model.WebPart) {
        jQuery.ajax({
            url: String['format']("{0}/_api/web/webs", _spPageContextInfo.webAbsoluteUrl),
            type: 'get',
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            success: function (d) {
                var stringBuilder = [];
                stringBuilder.push("<ul>");
                jQuery.each(d.value, function (id, val) {
                    stringBuilder.push(String['format']("<li>{0}</li>", val.Title));
                });
                stringBuilder.push("<ul>");
                webpart.instance.html(stringBuilder.join(''));
            }
        });
    }

    /* HelloWorld */
    /* REQUIRED: None */
    export function HelloWorldWebPart(webpart: Model.WebPart) {
        webpart.instance.html("<p>Hello world</p>");
    }

    /* Permissions */
    /* REQUIRED: [List], [ItemsCount] */
    export function ListItems(webpart: Model.WebPart) {
        var properties = webpart.properties[0];

        jQuery.ajax({
            url: String['format']("{0}/_api/web/lists/getByTitle('{1}')/items?$top={2}", _spPageContextInfo.webAbsoluteUrl, properties["List"], properties["ItemsCount"]),
            type: 'get',
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            success: function (d) {
                var stringBuilder = [];
                stringBuilder.push("<ul>");
                jQuery.each(d.value, function (id, val) {
                    stringBuilder.push(String['format']("<li>{0}</li>", val.Title));
                });
                stringBuilder.push("<ul>");
                webpart.instance.html(stringBuilder.join(''));
            }
        });
    }

    /* Permissions */
    /* REQUIRED: None */
    export function Permissions(webpart: Model.WebPart) {
        var properties = webpart.properties[0];

        jQuery.ajax({
            url: String['format']("{0}/_api/web/RoleAssignments?$expand=Member,Member/Users,RoleDefinitionBindings", _spPageContextInfo.webAbsoluteUrl),
            type: 'get',
            headers: {
                'accept': 'application/json;odata=verbose'
            },
            success: function (d) {
                var stringBuilder = [];
                stringBuilder.push("<ul>");
                jQuery.each(d.d.results, function (id, val) {
                    stringBuilder.push(String['format']("<li>'{0}' has permission '{1}' on this web.</li>", val.Member.Title, val.RoleDefinitionBindings.results[0].Name));

                    if (val.Member.Users && val.Member.Users.results && val.Member.Users.results.length > 0) {
                        jQuery.each(val.Member.Users.results, function (id, val) {
                            var externalUser = (val.LoginName.indexOf("#ext#") != -1) ? " <b>and is an external user</b>" : "";
                            stringBuilder.push(String['format']("<li style='margin-left: 40px;'><i>{0} is a member of this group{1}.</i></li>", val.Title, externalUser));
                        });
                    }
                });
                stringBuilder.push("<ul>");
                webpart.instance.html(stringBuilder.join(''));
            }
        });
    }
    export function ExternalUsers(webpart: Model.WebPart) {
        var properties = webpart.properties[0];

        jQuery.ajax({
            url: String['format']("{0}/_api/web/RoleAssignments?$expand=Member,Member/Users,RoleDefinitionBindings", _spPageContextInfo.webAbsoluteUrl),
            type: 'get',
            headers: {
                'accept': 'application/json;odata=verbose'
            },
            success: function (d) {
                var stringBuilder = [];
                stringBuilder.push("<ul>");
                jQuery.each(d.d.results, function (id, val) {
                    if (val.Member.LoginName.indexOf("#ext#") != -1) {
                        stringBuilder.push(String['format']("<li>{0}</li>", val.Member.Title));
                    }

                    if (val.Member.Users && val.Member.Users.results && val.Member.Users.results.length > 0) {
                        jQuery.each(val.Member.Users.results, function (id, user) {
                            if (user.LoginName.indexOf("#ext#") != -1) {
                                stringBuilder.push(String['format']("<li>{0} has {1}</li>", user.Title, val.RoleDefinitionBindings.results[0].Name));
                            }
                        });
                    }
                });
                stringBuilder.push("<ul>");
                webpart.instance.html(stringBuilder.join(''));
            }
        });
    }
    /* PermissionsWithCSS */
    /* REQUIRED: [CSS] */
    export function PermissionsWithCSS(webpart: Model.WebPart) {
        var properties = webpart.properties[0];

        jQuery("head").append(String['format']("<link rel='stylesheet' type='text/css' href='{0}' />", webpart.properties[0]["CSS"]));
        jQuery.ajax({
            url: String['format']("{0}/_api/web/RoleAssignments?$expand=Member,Member/Users,RoleDefinitionBindings", _spPageContextInfo.webAbsoluteUrl),
            type: 'get',
            headers: {
                'accept': 'application/json;odata=verbose'
            },
            success: function (d) {
                var stringBuilder = [];
                stringBuilder.push("<ul id='permissions-web'>");

                var extUsers = 0;

                jQuery.each(d.d.results, function (id, val) {
                    stringBuilder.push(String['format']("<li><b>{0}</b> has permission <b>{1}</b> on this web.</li>", val.Member.Title, val.RoleDefinitionBindings.results[0].Name));

                    if (val.Member.Users && val.Member.Users.results && val.Member.Users.results.length > 0) {
                        stringBuilder.push("<li style='margin-left: 40px;'><b>Members:</b></li>");
                        jQuery.each(val.Member.Users.results, function (id, val) {
                            var externalUser = (val.LoginName.indexOf("#ext#") != -1) ? " <b>(external user)</b>" : "";
                            stringBuilder.push(String['format']("<li style='margin-left: 40px;'><i>{0}{1}.</i></li>", val.Title, externalUser));

                            if (externalUser != "") {
                                extUsers++;
                            }
                        });
                    }
                });
                stringBuilder.push("</ul>");
                stringBuilder.push(String['format']("<p><b>A total of {0} external users has access to this site.</b></p>", extUsers));
                webpart.instance.html(stringBuilder.join(''));
            }
        });
    }

    /* CalendarItems */
    /* REQUIRED: [Web], [List], [ItemsCount], [Category] */
    export function CalendarItems(webpart: Model.WebPart) {
        var properties = webpart.properties[0];

        jQuery.ajax({
            url: String['format']("{0}{1}/_api/web/lists/getByTitle('{2}')/items?$top={3}&$filter=Category eq '{4}'", _spPageContextInfo.webAbsoluteUrl, properties["Web"], properties["List"], properties["ItemsCount"], properties["Category"]),
            type: 'get',
            headers: {
                'accept': 'application/json;odata=nometadata'
            },
            success: function (d) {
                var stringBuilder = [];
                stringBuilder.push("<ul>");
                jQuery.each(d.value, function (id, val) {
                    stringBuilder.push(String['format']("<li>{0}</li>", val.Title));
                });
                stringBuilder.push("<ul>");
                webpart.instance.html(stringBuilder.join(''));
            }
        });
    }

    module Util {
        export function search(queryText, managedProperties, rowLimit, startRow, allResults) {
            var allResults = allResults || [];
            var url = String['format']("{0}/_api/search/query?querytext='{1}'&rowlimit={2}&startrow={3}&selectproperties='{4}'&trimduplicates=false", _spPageContextInfo.webAbsoluteUrl, queryText, rowLimit, startRow, managedProperties);
            return $.getJSON(url).then(function(data) {
                var relevantResults = data.PrimaryQueryResult.RelevantResults;
                allResults = allResults.concat(relevantResults.Table.Rows);
                if (relevantResults.TotalRows > startRow + relevantResults.RowCount) {
                    return search(queryText, managedProperties, rowLimit, startRow + relevantResults.RowCount, allResults);
                }
                return allResults;
            });
        }
    }
}