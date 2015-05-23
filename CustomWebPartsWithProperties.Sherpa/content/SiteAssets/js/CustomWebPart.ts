﻿/// <reference path="SharePoint.d.ts" />
/// <reference path="jquery.d.ts" />

module CustomWebPart {
    module Util {
        export function GetWebPartsDefinitions() {
            return jQuery(Properties.WebPartClass);
        }
        export function ReplaceAll(str: string, f: string, r: string) {
            return str ? str.replace(new RegExp(f, '\g'), r) : "";
        }
        export function GetToolPaneForWebPart(webpartid: string) {
            return jQuery(".ms-TPBody[id*='" + ReplaceAll(webpartid, '-', '_') + "']").first();
        }
        export function GetHiddenInputFieldForWebPart(webpartid: string) {
            return jQuery(".aspNetHidden input[name*='" + webpartid + "']");
        }
        function GetUpdatedWebPartHtml(instance: any) {
            var properties = instance.data("webpart-properties")[0];

            for (var i = 0; i < Object.keys(properties).length; i++) {
                var key = Object.keys(properties)[i];
                var $input = jQuery("input.UserInput[name*='EditorZone'][name*='" + key + "']");
                properties[key] = $input.val();
            }
            instance.attr("data-webpart-properties", "[" + JSON.stringify(properties) + "]");

            return $('<div>').append(instance.clone()).html();
        }
        export function Log(message: string) {
            console.info(message);
        }
        export function Error(message: string) {
            console.error(message);
        }
        export function InEditMode() {
            var formName = (typeof window['MSOWebPartPageFormName'] === "string") ? window['MSOWebPartPageFormName'] : "aspnetForm";
            var form = window.document.forms[formName];

            if (form && ((form['MSOLayout_InDesignMode'] && form['MSOLayout_InDesignMode'].value) || (typeof window['MSOLayout_IsWikiEditMode'] === "function" && window['MSOLayout_IsWikiEditMode']()))) {
                return true;
            } else {
                return false;
            }
        }
        export function RenderWebPartProperties(webpart: Model.WebPart) {
            var $toolPane = GetToolPaneForWebPart(webpart.id[1]);
            if (Object.keys(webpart.properties[0]).length > 0 && $toolPane.length > 0) {
                jQuery(".ms-rte-embedcode-linkedit").hide();
                jQuery.getJSON(_spPageContextInfo.siteAbsoluteUrl + Properties.HtmlRootPath + "customproperties.txt", function (d) {
                    var props = [];

                    for (var i = 0; i < Object.keys(webpart.properties[0]).length; i++) {
                        var key = Object.keys(webpart.properties[0])[i];
                        var value = webpart.properties[0][key];

                        props.push(String['format'](d.Field, Util.ReplaceAll(webpart.id[1], '-', '_'), key, value))
                    }

                    $toolPane.append(String['format'](d.Container, Util.ReplaceAll(webpart.id[1], '-', '_'), props.join('')));

                    var $submit = jQuery("input[type='submit'][name*='OKBtn'], input[type='submit'][name*='AppBtn']");
                    $submit.click(function (event, args) {
                        GetHiddenInputFieldForWebPart(webpart.id[1]).val(GetUpdatedWebPartHtml(webpart.instance));
                    });
                });
            }
        }
    }
    export module Properties {
        export var WebPartClass = '.custom-webpart';
        export var HtmlRootPath = "/siteassets/customwebparts/html/";
    }
    export module Model {
        export class WebPart {
            instance: any;
            id: Array<string>;
            renderfunction: string;
            properties: Array<Object>;
            render() {
                Manager.Render(this);
            }
            move(zoneID : string, zoneIndex : number) {
                Manager.Move(this, zoneID, zoneIndex);
            }
            delete() {
                Manager.Delete(this);
            }

            constructor(element: any) {
                this.instance = element;
                this.id = [
                    this.instance.parents("div[webpartid]").first().attr("webpartid"),
                    this.instance.parents("div[webpartid2]").first().length > 0 ? this.instance.parents("div[webpartid2]").first().attr("webpartid2") : this.instance.parents("div[webpartid]").first().attr("webpartid")
                ]
                this.renderfunction = this.instance.data("webpart-renderfunction");
                this.properties = this.instance.data("webpart-properties");
            }
        }
        export var WebParts = [];
    }
    export module Manager {
        export function Init() {
            Util.GetWebPartsDefinitions().each(function () {
                try {
                    Model.WebParts.push(new Model.WebPart(jQuery(this)));
                } catch (e) {
                    Util.Error("Error parsing webpart.");
                }
            });

            RenderAllWebParts();
        }
        function RenderAllWebParts() {
            for (var i in Model.WebParts) {
                Model.WebParts[i].render();
            }
        }
        export function DeleteAllWebParts() {
            for (var i in Model.WebParts) {
                Model.WebParts[i].delete();
            }
            window.setTimeout(function () { location.href = location.href; }, 3000);
        }
        export function Render(webpart: Model.WebPart) {
            if (!Util.InEditMode()) {
                try {
                    eval(webpart.renderfunction + "(webpart)");
                } catch (e) {
                    Util.Error("The render function for one of the webparts doesn't exist");
                }
            } else {
                Util.RenderWebPartProperties(webpart);
            }
        }
        export function Delete(webpart: Model.WebPart) {
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
        export function Move(webpart: Model.WebPart, zoneID : string, zoneIndex : number) {
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
    }
}


jQuery(function () {
    CustomWebPart.Manager.Init();
});