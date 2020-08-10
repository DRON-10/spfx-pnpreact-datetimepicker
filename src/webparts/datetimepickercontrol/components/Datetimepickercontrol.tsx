import * as React from 'react';  
import styles from './Datetimepickercontrol.module.scss';  
import { IDatetimepickercontrolProps } from './IDatetimepickercontrolProps';  
import { IDatetimepickercontrolState } from './IDatetimepickercontrolState';  
import { TextField } from 'office-ui-fabric-react/lib/TextField';  
import { MessageBar, MessageBarType, IStackProps, Stack } from 'office-ui-fabric-react';  
import { autobind } from 'office-ui-fabric-react';  
import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/dateTimePicker';  
import { sp } from "@pnp/sp";  
import "@pnp/sp/webs";  
import "@pnp/sp/lists";  
import "@pnp/sp/items";  
  
const verticalStackProps: IStackProps = {  
  styles: { root: { overflow: 'hidden', width: '100%' } },  
  tokens: { childrenGap: 20 }  
};  
  
export default class Datetimepickercontrol extends React.Component<IDatetimepickercontrolProps, IDatetimepickercontrolState> {  
  constructor(props: IDatetimepickercontrolProps, state: IDatetimepickercontrolState) {  
    super(props);  
    sp.setup({  
      spfxContext: this.props.context  
    });  
    this.state = {  
      projectTitle: '',  
      projectDescription: '',  
      startDate: new Date(),  
      endDate: new Date(),  
      showMessageBar: false  
    };  
  }  
  
  public render(): React.ReactElement<IDatetimepickercontrolProps> {  
    return (  
  
      <div className={styles.row}>  
        <h1>Create New Project</h1>  
        {  
          this.state.showMessageBar  
            ?  
            <div className="form-group">  
              <Stack {...verticalStackProps}>  
                <MessageBar messageBarType={this.state.messageType}>{this.state.message}</MessageBar>  
              </Stack>  
            </div>  
            :  
            null  
        }  
        <div className={styles.row}>  
          <TextField label="Project Title" required onChanged={this.__onchangedTitle} />  
          <TextField label="Project Description" required onChanged={this.__onchangedDescription} />  
          <DateTimePicker label="Start Date"  
            dateConvention={DateConvention.DateTime}  
            timeConvention={TimeConvention.Hours12}  
            timeDisplayControlType={TimeDisplayControlType.Dropdown}  
            showLabels={false}  
            value={this.state.startDate}  
            onChange={this.__onchangedStartDate}  
          />  
          <DateTimePicker label="End Date"  
            dateConvention={DateConvention.Date}  
            timeConvention={TimeConvention.Hours12}  
            timeDisplayControlType={TimeDisplayControlType.Dropdown}  
            showLabels={false}  
            value={this.state.endDate}  
            onChange={this.__onchangedEndDate}  
          />  
          <div className={styles.button}>  
            <button type="button" className="btn btn-primary" onClick={this.__createItem}>Submit</button>  
          </div>  
        </div>  
      </div>  
    );  
  }  
  @autobind  
  private __onchangedTitle(Title: any): void {  
    this.setState({ projectTitle: Title });  
  }  
  
  @autobind  
  private __onchangedDescription(description: any): void {  
    this.setState({ projectDescription: description });  
  }  
  
  @autobind  
  private __onchangedStartDate(DateFrom: any): void {  
    this.setState({ startDate: DateFrom });  
  }  
  
  @autobind  
  private __onchangedEndDate(DateTo: any): void {  
    this.setState({ endDate: DateTo });  
  }  
  
  @autobind  
  private async __createItem() {  
    try {  
      await sp.web.lists.getByTitle('DateRangeList').items.add({  
        Title: this.state.projectTitle,  
        description: this.state.projectDescription,  
        DateFrom: this.state.startDate,  
        DateTo: this.state.endDate  
      });  
      this.setState({  
        message: "Item: " + this.state.projectTitle + " - created successfully!",  
        showMessageBar: true,  
        messageType: MessageBarType.success  
      });  
    }  
    catch (error) {  
      this.setState({  
        message: "Item " + this.state.projectTitle + " creation failed with error: " + error,  
        showMessageBar: true,  
        messageType: MessageBarType.error  
      });  
    }  
  }  
} 