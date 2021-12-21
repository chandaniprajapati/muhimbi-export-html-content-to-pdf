
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from "@microsoft/sp-webpart-base";

declare global {
    interface Navigator {
        msSaveBlob?: (blob: any, defaultName?: string) => boolean;
        msSaveOrOpenBlob: (blob: Blob) => void
    }
}

export class ConvertFileService {

    constructor(private context: WebPartContext) {

    }

    public generateBlob(base64string) {
        var file_bytes = atob(base64string);
        var byte_numbers = new Array(file_bytes.length);
        for (var i = 0; i < file_bytes.length; i++) {
            byte_numbers[i] = file_bytes.charCodeAt(i);
        }
        var byte_array = new Uint8Array(byte_numbers);
        var file_blob = new Blob([byte_array], { type: "application/pdf" });
        return file_blob;
    }

    public async downloadFile(API_Key: string, url: string, fileName: string, request: any, extension: string) {
        if (API_Key == '')
            return;

        try {
            var input_data = JSON.stringify(request);
            const requestHeaders: Headers = new Headers();
            requestHeaders.append('Content-type', 'application/json');
            requestHeaders.append('API_key', API_Key);
            const httpClientOptions: IHttpClientOptions = {
                headers: requestHeaders,
                body: input_data
            };
            return this.context.httpClient.post(url, HttpClient.configurations.v1, httpClientOptions)
                .then(async (response: HttpClientResponse) => {
                    let data = await response.json();
                    if (data['result_code'] == "Success") {
                        var file_blob = this.generateBlob(data['processed_file_content']);
                        if ((window.navigator as any).msSaveBlob) {
                            (window.navigator as any).msSaveOrOpenBlob(file_blob, data['base_file_name'] + extension);
                        }
                        else {
                            var download_link = window.document.createElement("a");
                            download_link.href = window.URL.createObjectURL(file_blob);
                            download_link.download = fileName;
                            download_link.click();
                        }
                    }
                    return data;
                }).catch((error: any): Promise<any> => {
                    console.log("Getting error while dowloading file:", error);
                    return error;
                });
        }
        catch (error) {
            console.log(error.message);
        }
    }

    public convertToPDF(API_Key: string, apiUrl: string, htmlContent: string, fileName: string) {
        const base64data = btoa(htmlContent);
        const postURL = apiUrl;
        const body = {          
            "use_async_pattern": false,
			"source_file_name": `${fileName}.html`,
			"source_file_content": base64data,
			"output_format": "PDF",
        };
        return this.downloadFile(API_Key, postURL, fileName, body, '.pdf');
    }
}