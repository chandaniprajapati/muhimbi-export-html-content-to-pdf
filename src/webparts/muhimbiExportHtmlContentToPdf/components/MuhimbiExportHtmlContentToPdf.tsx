import * as React from 'react';
import styles from './MuhimbiExportHtmlContentToPdf.module.scss';
import { IMuhimbiExportHtmlContentToPdfProps } from './IMuhimbiExportHtmlContentToPdfProps';
import { IMuhimbiExportHtmlContentToPdfState } from './IMuhimbiExportHtmlContentToPdfState';
import { ConvertFileService } from '../../../services/ConvertFileService';
import { PrimaryButton, Stack } from 'office-ui-fabric-react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { escape } from '@microsoft/sp-lodash-subset';

export default class MuhimbiExportHtmlContentToPdf extends React.Component<IMuhimbiExportHtmlContentToPdfProps, IMuhimbiExportHtmlContentToPdfState> {

  private _services: ConvertFileService = null;

  private htmlContent = `<table style="width:100%;border: 1px solid black;border-collapse: collapse;"">
                          <tr>
                            <th style="border: 1px solid black;border-collapse: collapse;padding: 5px;">Firstname</th>
                            <th style="border: 1px solid black;border-collapse: collapse;padding: 5px;">Lastname</th> 
                            <th style="border: 1px solid black;border-collapse: collapse;padding: 5px;">Age</th>
                          </tr>
                          <tr>
                            <td style="border: 1px solid black;border-collapse: collapse;padding: 5px;">John</td>
                            <td style="border: 1px solid black;border-collapse: collapse;padding: 5px;">Smith</td>
                            <td style="border: 1px solid black;border-collapse: collapse;padding: 5px;">50</td>
                          </tr>
                          <tr>
                            <td style="border: 1px solid black;border-collapse: collapse;padding: 5px;">Eve</td>
                            <td style="border: 1px solid black;border-collapse: collapse;padding: 5px;">Jackson</td>
                            <td style="border: 1px solid black;border-collapse: collapse;padding: 5px;">94</td>
                          </tr>
                          <tr>
                            <td style="border: 1px solid black;border-collapse: collapse;padding: 5px;">John</td>
                            <td style="border: 1px solid black;border-collapse: collapse;padding: 5px;">Doe</td>
                            <td style="border: 1px solid black;border-collapse: collapse;padding: 5px;">80</td>
                          </tr>
                          </table>`;

  constructor(props: IMuhimbiExportHtmlContentToPdfProps) {
    super(props);
    this.state = {
      showLoader: false
    }
    this._services = new ConvertFileService(this.props.context);
  }

  public exportToPDF = () => {
    this.setState({ showLoader: true });
    this._services.convertToPDF(this.props.apiKey, this.props.apiUrl, this.htmlContent, 'test.pdf').then(result => {
      this.setState({ showLoader: false });
    }).catch(error => {
      console.log("Getting error in export to PDF:", error);
      this.setState({ showLoader: false });
    });
  }

  public render(): React.ReactElement<IMuhimbiExportHtmlContentToPdfProps> {

    const { description } = this.props;
    const { showLoader } = this.state;

    return (
      <div className={styles.muhimbiExportHtmlContentToPdf}>
        <h1>{description}</h1>
        <Stack>
          {showLoader &&
            <Spinner label="Wait, wait..., file is downloading" size={SpinnerSize.large} />
          }
          <PrimaryButton disabled={showLoader ? true : false} style={{ width: '200px', display: 'flex', alignSelf: 'end', marginBottom: '15px' }} text="Download as PDF" onClick={this.exportToPDF} allowDisabledFocus />
          <div dangerouslySetInnerHTML={{ __html: this.htmlContent }} />
        </Stack>
      </div>
    );
  }
}
