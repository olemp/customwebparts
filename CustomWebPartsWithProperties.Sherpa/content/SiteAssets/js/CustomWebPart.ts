/// <reference path="SharePoint.d.ts" />
/// <reference path="jquery.d.ts" />

module CustomWebPart {
    export module Properties {
        export var WebPartClass = '.custom-webpart';
        export var HtmlRootPath = "/siteassets/customwebparts/html/";
    }
    export module Model {
        export class WebPart {
            instance: any;
            id: string;
            renderfunction: string;
            properties: Array<Object>;
            render() {
                Manager.Render(this);
            }

            constructor(element: any) {
                this.instance = element;
                this.id = this.instance.closest("div[webpartid2]").attr("webpartid2");
                this.renderfunction = this.instance.data("webpart-renderfunction");
                this.properties = this.instance.data("webpart-properties");
            }
        }
        export var WebParts = [];
    }
    export module Manager {
        module Util {
            export function GetWebPartsDefinitions() {
                return jQuery(Properties.WebPartClass);
            }
            export function ReplaceAll(str : string, f : string, r : string) {
                return str.replace(new RegExp(f, '\g'), r)
            }
            export function GetToolPaneForWebPart(webpartid : string) {
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
                var $toolPane = GetToolPaneForWebPart(webpart.id);
                if ($toolPane.length > 0) {
                    jQuery(".ms-rte-embedcode-linkedit").hide();
                    jQuery.getJSON(_spPageContextInfo.siteAbsoluteUrl + Properties.HtmlRootPath + "customproperties.txt", function (d) {
                        var props = [];

                        for (var i = 0; i < Object.keys(webpart.properties[0]).length; i++) {
                            var key = Object.keys(webpart.properties[0])[i];
                            var value = webpart.properties[0][key];

                            props.push(String['format'](d.Field, Util.ReplaceAll(webpart.id, '-', '_'), key, value))
                        }

                        $toolPane.append(String['format'](d.Container, Util.ReplaceAll(webpart.id, '-', '_'), props.join('')));

                        var $submit = jQuery("input[type='submit'][name*='OKBtn']");
                        $submit.click(function (event, args) {
                            GetHiddenInputFieldForWebPart(webpart.id).val(GetUpdatedWebPartHtml(webpart.instance));
                        });
                    });
                }
            }
        }
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
    }
}


jQuery(function () {
    CustomWebPart.Manager.Init();
});