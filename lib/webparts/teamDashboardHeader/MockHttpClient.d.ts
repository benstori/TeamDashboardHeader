import { ISPList } from './TeamDashboardHeaderWebPart';
export default class MockHttpClient {
    private static _items;
    static get(): Promise<ISPList[]>;
}
