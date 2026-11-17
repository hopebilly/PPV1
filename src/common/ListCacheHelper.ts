import axios, { AxiosRequestConfig } from "axios";
import moment from 'moment';

export class ListCacheHelper {
  private async _getListsLastModifiedDate(dataSiteUrl: string, listName: string): Promise<Date> {
    const config: AxiosRequestConfig = {
      withCredentials: true,
      method: "GET",
      headers: {
        accept: "application/json;odata=verbose", //It defines the Data format
        "content-type": "application/json;odata=verbose", //It defines the content type as JSON
        "Cache-Control": "no-cache",   // Prevent caching by the browser
        "Pragma": "no-cache",          // Prevent caching in older HTTP/1.0 proxies
        "Expires": "0",                // Ensures no caching (Expires header for proxies)
      },
    };

    const response = await axios.get(`${dataSiteUrl}/_api/web/lists/getbytitle('${listName}')/LastItemModifiedDate`, config,
                                  
    );
    const listModifiedDate: Date = response.data.d.LastItemModifiedDate as Date;

    return listModifiedDate;
  }

  public async DoesListValueNeedRefresh(dataSiteUrl: string, listName: string, lastCached?: Date): Promise<boolean> {
    const responseForDate = await this._getListsLastModifiedDate(dataSiteUrl, listName);
    const momentLastModified = moment(responseForDate);

    let needRefresh: boolean = false;
    if (lastCached) {
      const cachedOnDate: Date = lastCached as Date;
      const momemntCachedOnDate = moment(cachedOnDate);
      console.log(`CURRENT LIST VALUES for ${listName} CACHED ON: ${momemntCachedOnDate.toString()}`);
      const duration = moment.duration(momentLastModified.diff(momemntCachedOnDate)).milliseconds();
      console.log("Duration: " + duration);
      if (duration <= 0) {
        /*last modified date is less or equal to cached date */
        console.log(`USED CACHED LIST VALUES for ${listName}: ${moment(cachedOnDate).toString()}`);
      } else {
        needRefresh = true;
      }
    } else {
      needRefresh = true;
    }

    return needRefresh;
  }
}