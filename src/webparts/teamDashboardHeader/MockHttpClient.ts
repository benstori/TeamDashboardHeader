import { ISPList } from './TeamDashboardHeaderWebPart';

export default class MockHttpClient  {

   private static _items: ISPList[] = [{ Title: 'Mock List', Id: '1', DeptURL:'/gsrv_teams/sdgdev' }];

   public static get(): Promise<ISPList[]> {
   return new Promise<ISPList[]>((resolve) => {
           resolve(MockHttpClient._items);
       });
   }
}