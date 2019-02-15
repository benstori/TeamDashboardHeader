import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';
export interface ITeamDashboardHeaderWebPartProps {
    description: string;
}
export interface ISPLists {
    value: ISPList[];
}
export interface ISPList {
    Title: string;
    Id: string;
    DeptURL: string;
}
export default class TeamDashboardHeaderWebPart extends BaseClientSideWebPart<ITeamDashboardHeaderWebPartProps> {
    getuser: Promise<{}>;
    render(): void;
    protected readonly dataVersion: Version;
    private _getMockListData();
    _getListData(): Promise<ISPLists>;
    private _renderListAsync();
    onInit(): Promise<void>;
    private _renderList(items);
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
