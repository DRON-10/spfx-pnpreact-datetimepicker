import { MessageBarType } from 'office-ui-fabric-react';   
  
export interface IDatetimepickercontrolState{  
    First_Name: string;  
    Last_Name: string;  
    startDate: Date;  
    endDate: Date;  
    showMessageBar: boolean;      
    messageType?: MessageBarType;      
    message?: string;    
}