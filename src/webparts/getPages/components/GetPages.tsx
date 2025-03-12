import * as React from 'react';
import type { IGetPagesProps } from './IGetPagesProps';
import { PrimaryButton } from '@fluentui/react';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

const GetPages: React.FunctionComponent<IGetPagesProps> = (props: IGetPagesProps) => {

  let addedPageUrl: boolean = false;
  const pageLimit = 20;
  const [pages, setPages] = React.useState<string[]>([`https://my-gsk.atlassian.net/wiki/rest/api/search?start=0&limit=${pageLimit}&cql=type=page&expand=title,content.body.view`]);
  const [pageIndex, setPageIndex] = React.useState<number>(0);
  const [hasNext, setHasNext] = React.useState<boolean>(true);
  
  const getPages = async (): Promise<void> => {
    props.context.aadHttpClientFactory.getClient('https://service.flow.microsoft.com/').then((client: AadHttpClient) => {
      client.post("https://prod-157.westus.logic.azure.com:443/workflows/bf303da5967a472a8f69d511ac8925d6/triggers/manual/paths/invoke?api-version=2016-06-01", AadHttpClient.configurations.v1, {
        headers: { 'Content-type': 'application/json' },
        body: JSON.stringify({ "url": pages[pageIndex] })
      }).then((resp: HttpClientResponse): Promise<any> => {
        console.log("resp", resp);
        return resp.json();
      }).then((data: any) => {
        console.log("data", data);
        if (typeof data._links.next !== typeof undefined && !addedPageUrl) {
          if (pages.length - 1 === pageIndex && !pages.some(p => p === `${data._links.base}${data._links.next}`)) {
            setPages(pages => [...pages, `${data._links.base}${data._links.next}`]);
            addedPageUrl = true;
          }
          else {
            setHasNext(false);
          }
        }
      });
    });
  }

  React.useEffect(() => {
    getPages();
  }, []);

  React.useEffect(() => {
    addedPageUrl = false;
    getPages();
  }, [pageIndex]);
  React.useEffect(() => { console.log("pages", pages); }, [pages]);

  return (
    <div>
      <PrimaryButton
        text='Prev'
        disabled={pageIndex === 0}
        onClick={() => {
          const newPageIndex = pageIndex - 1;
          setPageIndex(newPageIndex);
        }}
      />
      <PrimaryButton
        text='Next'
        disabled={!hasNext}
        onClick={() => {
          const newPageIndex = pageIndex + 1;
          setPageIndex(newPageIndex);
        }}
      />
    </div>
  );
}

export default GetPages;