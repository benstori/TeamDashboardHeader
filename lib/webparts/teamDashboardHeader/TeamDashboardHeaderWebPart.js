"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_1 = require("@pnp/sp");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var sp_core_library_2 = require("@microsoft/sp-core-library");
var sp_http_1 = require("@microsoft/sp-http");
var TeamDashboardHeaderWebPart_module_scss_1 = require("./TeamDashboardHeaderWebPart.module.scss");
var strings = require("TeamDashboardHeaderWebPartStrings");
var MockHttpClient_1 = require("./MockHttpClient");
var logo = require('./assests/Team.png');
//global vars
var userDept = "";
var TeamDashboardHeaderWebPart = (function (_super) {
    __extends(TeamDashboardHeaderWebPart, _super);
    function TeamDashboardHeaderWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.getuser = new Promise(function (resolve, reject) {
            // SharePoint PnP Rest Call to get the User Profile Properties
            return sp_1.sp.profiles.myProperties.get().then(function (result) {
                var props = result.UserProfileProperties;
                var propValue = "";
                var userDepartment = "";
                props.forEach(function (prop) {
                    //this call returns key/value pairs so we need to look for the Dept Key
                    if (prop.Key == "Department") {
                        // set our global var for the users Dept.
                        userDept += prop.Value;
                    }
                });
                return result;
            }).then(function (result) {
                _this._getListData().then(function (response) {
                    _this._renderList(response.value);
                });
            });
        });
        return _this;
    }
    TeamDashboardHeaderWebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n      <div class=\"" + TeamDashboardHeaderWebPart_module_scss_1.default.teamDashboardHeader + "\">\n                <div id=\"TeamDashboardHeader\"/>\n      </div>";
        //this._renderListAsync();
    };
    Object.defineProperty(TeamDashboardHeaderWebPart.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    TeamDashboardHeaderWebPart.prototype._getMockListData = function () {
        return MockHttpClient_1.default.get()
            .then(function (data) {
            var listData = { value: data };
            return listData;
        });
    };
    // main REST Call to the list...passing in the deaprtment into the call to 
    //return a single list item
    TeamDashboardHeaderWebPart.prototype._getListData = function () {
        return this.context.spHttpClient.get("https://girlscoutsrv.sharepoint.com/_api/web/lists/GetByTitle('TeamDashboardSettings')/Items?$filter=Title eq '" + userDept + "'", sp_http_1.SPHttpClient.configurations.v1)
            .then(function (response) {
            return response.json();
        });
    };
    //mock up 
    TeamDashboardHeaderWebPart.prototype._renderListAsync = function () {
        var _this = this;
        // Local environment
        if (sp_core_library_2.Environment.type === sp_core_library_2.EnvironmentType.Local) {
            this._getMockListData().then(function (response) {
                _this._renderList(response.value);
            });
        }
        else if (sp_core_library_2.Environment.type == sp_core_library_2.EnvironmentType.SharePoint ||
            sp_core_library_2.Environment.type == sp_core_library_2.EnvironmentType.ClassicSharePoint) {
            this._getListData()
                .then(function (response) {
                _this._renderList(response.value);
            });
        }
    };
    // this is required to use the SharePoint PnP shorthand REST CALLS
    TeamDashboardHeaderWebPart.prototype.onInit = function () {
        var _this = this;
        return _super.prototype.onInit.call(this).then(function (_) {
            sp_1.sp.setup({
                spfxContext: _this.context
            });
        });
    };
    TeamDashboardHeaderWebPart.prototype._renderList = function (items) {
        var html = '';
        items.forEach(function (item) {
            html += "\n      <table style=\"width:100%;height:1px;\">\n        <tr>\n          <td style=\"height:1px;text-align:right;width:12%\">\n          <a href=\"" + item.DeptURL + " target=\"_blank\">\n             <img id=\"TeamImage\" class=\"" + TeamDashboardHeaderWebPart_module_scss_1.default.headerImage + "\" src=\"" + logo + "\" alt=\"GSVR Logo\" /></a>\n          </td>\n          <td class=\"width:70%;height:1px;vertical-align:middle;\"> \n          <h2 class=\"" + TeamDashboardHeaderWebPart_module_scss_1.default.h2 + "\"><a id=\"teamHeaderLink\" href=\"" + item.DeptURL + "\" target=\"_blank\">Team Dashboard</a></h2>\n          </td>\n        </tr>\n      </table>   \n        ";
        });
        var listContainer = this.domElement.querySelector('#TeamDashboardHeader');
        listContainer.innerHTML = html;
    };
    TeamDashboardHeaderWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return TeamDashboardHeaderWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports.default = TeamDashboardHeaderWebPart;

//# sourceMappingURL=TeamDashboardHeaderWebPart.js.map
