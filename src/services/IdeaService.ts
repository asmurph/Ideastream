import { IIdeaListItem } from "../models";
import { 
    SPHttpClient,
    SPHttpClientResponse,
    ISPHttpClientOptions,
} from '@microsoft/sp-http';

const LIST_API_ENDPOINT: string = `/_api/web/lists/getbytitle('Ideas')`;
const SELECT_QUERY: string = `$select=Id,Title,Description,IdeaImage,Created,Comments`;

export class IdeaService {

    private _spHttpOptions: any = {
        getNoMetaData: <ISPHttpClientOptions> {
            headers: {'ACCEPT':'application/json; odata.metadata=none' }
        }
    };

    constructor(private siteAbsoluteUrl: string, private client: SPHttpClient){

    }
    /* Returns a collection of all ideas */

    public getIdeas(sortOrder?: string): Promise<IIdeaListItem[]> {
        let promise: Promise<IIdeaListItem[]> = new Promise<IIdeaListItem[]>((resolve, reject) => {
            let query: string = `${this.siteAbsoluteUrl}${LIST_API_ENDPOINT}/items?${SELECT_QUERY}&${sortOrder}`;
            this.client.get(
                query,
                SPHttpClient.configurations.v1,
                this._spHttpOptions.getNoMetaData
            )
            .then((response: SPHttpClientResponse): Promise<{value: IIdeaListItem[]}> => {
                return response.json();
            })
            .then((response: {value: IIdeaListItem[]}) => {
                resolve(response.value);
            })
            .catch((error:any) => {
                reject(error);
            });
        });

        return promise;
    }
}