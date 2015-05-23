/// <reference path="SharePoint.d.ts" />
/// <reference path="jquery.d.ts" />
var CustomWebPart;
(function (CustomWebPart) {
    var Util;
    (function (Util) {
        function GetWebPartsDefinitions() {
            return jQuery(Properties.WebPartClass);
        }
        Util.GetWebPartsDefinitions = GetWebPartsDefinitions;
        function ReplaceAll(str, f, r) {
            return str ? str.replace(new RegExp(f, '\g'), r) : "";
        }
        Util.ReplaceAll = ReplaceAll;
        function GetToolPaneForWebPart(webpartid) {
            return jQuery(".ms-TPBody[id*='" + ReplaceAll(webpartid, '-', '_') + "']").first();
        }
        Util.GetToolPaneForWebPart = GetToolPaneForWebPart;
        function GetHiddenInputFieldForWebPart(webpartid) {
            return jQuery(".aspNetHidden input[name*='" + webpartid + "']");
        }
        Util.GetHiddenInputFieldForWebPart = GetHiddenInputFieldForWebPart;
        function GetUpdatedWebPartHtml(instance) {
            var properties = instance.data("webpart-properties")[0];
            for (var i = 0; i < Object.keys(properties).length; i++) {
                var key = Object.keys(properties)[i];
                var $input = jQuery("input.UserInput[name*='EditorZone'][name*='" + key + "']");
                properties[key] = $input.val();
            }
            instance.attr("data-webpart-properties", "[" + JSON.stringify(properties) + "]");
            return $('<div>').append(instance.clone()).html();
        }
        function Log(message) {
            console.info(message);
        }
        Util.Log = Log;
        function Error(message) {
            console.error(message);
        }
        Util.Error = Error;
        function InEditMode() {
            var formName = (typeof window['MSOWebPartPageFormName'] === "string") ? window['MSOWebPartPageFormName'] : "aspnetForm";
            var form = window.document.forms[formName];
            if (form && ((form['MSOLayout_InDesignMode'] && form['MSOLayout_InDesignMode'].value) || (typeof window['MSOLayout_IsWikiEditMode'] === "function" && window['MSOLayout_IsWikiEditMode']()))) {
                return true;
            }
            else {
                return false;
            }
        }
        Util.InEditMode = InEditMode;
        function RenderWebPartProperties(webpart) {
            var $toolPane = GetToolPaneForWebPart(webpart.id[1]);
            if (Object.keys(webpart.properties[0]).length > 0 && $toolPane.length > 0) {
                jQuery(".ms-rte-embedcode-linkedit").hide();
                jQuery.getJSON(_spPageContextInfo.siteAbsoluteUrl + Properties.HtmlRootPath + "customproperties.txt", function (d) {
                    var props = [];
                    for (var i = 0; i < Object.keys(webpart.properties[0]).length; i++) {
                        var key = Object.keys(webpart.properties[0])[i];
                        var value = webpart.properties[0][key];
                        props.push(String['format'](d.Field, Util.ReplaceAll(webpart.id[1], '-', '_'), key, value));
                    }
                    $toolPane.append(String['format'](d.Container, Util.ReplaceAll(webpart.id[1], '-', '_'), props.join('')));
                    var $submit = jQuery("input[type='submit'][name*='OKBtn'], input[type='submit'][name*='AppBtn']");
                    $submit.click(function (event, args) {
                        GetHiddenInputFieldForWebPart(webpart.id[1]).val(GetUpdatedWebPartHtml(webpart.instance));
                    });
                });
            }
        }
        Util.RenderWebPartProperties = RenderWebPartProperties;
    })(Util || (Util = {}));
    var Properties;
    (function (Properties) {
        Properties.WebPartClass = '.custom-webpart';
        Properties.HtmlRootPath = "/siteassets/customwebparts/html/";
    })(Properties = CustomWebPart.Properties || (CustomWebPart.Properties = {}));
    var Model;
    (function (Model) {
        var WebPart = (function () {
            function WebPart(element) {
                this.instance = element;
                this.id = [
                    this.instance.parents("div[webpartid]").first().attr("webpartid"),
                    this.instance.parents("div[webpartid2]").first().length > 0 ? this.instance.parents("div[webpartid2]").first().attr("webpartid2") : this.instance.parents("div[webpartid]").first().attr("webpartid")
                ];
                this.renderfunction = this.instance.data("webpart-renderfunction");
                this.properties = this.instance.data("webpart-properties");
            }
            WebPart.prototype.render = function () {
                Manager.Render(this);
            };
            WebPart.prototype.move = function (zoneID, zoneIndex) {
                Manager.Move(this, zoneID, zoneIndex);
            };
            WebPart.prototype.delete = function () {
                Manager.Delete(this);
            };
            return WebPart;
        })();
        Model.WebPart = WebPart;
        Model.WebParts = [];
    })(Model = CustomWebPart.Model || (CustomWebPart.Model = {}));
    var Manager;
    (function (Manager) {
        function Init() {
            Util.GetWebPartsDefinitions().each(function () {
                try {
                    Model.WebParts.push(new Model.WebPart(jQuery(this)));
                }
                catch (e) {
                    Util.Error("Error parsing webpart.");
                }
            });
            RenderAllWebParts();
        }
        Manager.Init = Init;
        function RenderAllWebParts() {
            for (var i in Model.WebParts) {
                Model.WebParts[i].render();
            }
        }
        function DeleteAllWebParts() {
            for (var i in Model.WebParts) {
                Model.WebParts[i].delete();
            }
            window.setTimeout(function () {
                location.href = location.href;
            }, 3000);
        }
        Manager.DeleteAllWebParts = DeleteAllWebParts;
        function Render(webpart) {
            if (!Util.InEditMode()) {
                try {
                    eval(webpart.renderfunction + "(webpart)");
                }
                catch (e) {
                    Util.Error("The render function for one of the webparts doesn't exist");
                }
            }
            else {
                Util.RenderWebPartProperties(webpart);
            }
        }
        Manager.Render = Render;
        function Delete(webpart) {
            var clientContext = new SP.ClientContext(_spPageContextInfo.webAbsoluteUrl);
            var oFile = clientContext.get_web().getFileByServerRelativeUrl(_spPageContextInfo.serverRequestPath);
            var limitedWebPartManager = oFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
            var webPartDefinition = limitedWebPartManager.get_webParts().getById(new SP.Guid(webpart.id[0]));
            webPartDefinition.deleteWebPart();
            clientContext.load(webPartDefinition);
            clientContext.executeQueryAsync(function () {
                console.log(String['format']("Webpart with ID '{0}' deleted.", webpart.id[0]));
            });
        }
        Manager.Delete = Delete;
        function Move(webpart, zoneID, zoneIndex) {
            var clientContext = new SP.ClientContext(_spPageContextInfo.webAbsoluteUrl);
            var oFile = clientContext.get_web().getFileByServerRelativeUrl(_spPageContextInfo.serverRequestPath);
            var limitedWebPartManager = oFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
            var webPartDefinition = limitedWebPartManager.get_webParts().getById(new SP.Guid(webpart.id[0]));
            webPartDefinition.moveWebPartTo(zoneID, zoneIndex);
            clientContext.load(webPartDefinition);
            clientContext.executeQueryAsync(function () {
                console.log(String['format']("Webpart with ID '{0}' moved to zone {1} with index {2}.", webpart.id[0], zoneID, zoneIndex));
            });
        }
        Manager.Move = Move;
    })(Manager = CustomWebPart.Manager || (CustomWebPart.Manager = {}));
})(CustomWebPart || (CustomWebPart = {}));
jQuery(function () {
    CustomWebPart.Manager.Init();
});
