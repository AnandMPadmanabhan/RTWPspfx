import * as React from 'react';
import styles from './RtwpForm.module.scss';
import { IRtwpFormProps } from './IRtwpFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as pnp from "sp-pnp-js";  
import { sp } from "@pnp/sp";  
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/items/list";
import "@pnp/sp/site-users/web";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import Form from 'react-bootstrap/Form';
import Button from 'react-bootstrap/Button';
import Row from 'react-bootstrap/Row';
import Col from 'react-bootstrap/Col';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
//import  Dropdown from 'react-bootstrap/Dropdown';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
//import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DefaultButton,PrimaryButton,PeoplePickerItem } from 'office-ui-fabric-react';
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
import 'D:/SPFx/Forms/RTWPForm/node_modules/bootstrap/dist/css/bootstrap.min.css';

var items1: IDropdownOption[]=[];
var items2: IDropdownOption[]=[];
var userId: Number;
var empName:string;
var selectedWork:number;
var selectedBuilding:number;
var AddInfo:string;
var disableButton:boolean=true;
var currUser :string="";
var days_selected:string[]=[];
var day_btncolor:string[]=["gainsboro","gainsboro","gainsboro","gainsboro","gainsboro","gainsboro","gainsboro"];
const stackTokens: IStackTokens = { childrenGap: 20 };
const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 }
};
export interface IReactGetItemsState{ 
  items1: IDropdownOption[];
  items2: IDropdownOption[];
  userId: Number;
  currUser :string;
  day_btncolor:string[];
  days_selected:string[];
  disableButton:boolean;
}
const Reason:IDropdownOption[]=[{key:"1",text:"Package Pickup"},{key:"2",text:"Project Requirements"}];
var reason_Selected:string="";
export default class RtwpForm extends React.Component<IRtwpFormProps, any> {
  public constructor(props: IRtwpFormProps, state: IReactGetItemsState){ 
    super(props); 
    this.state = { 
      items1:[],
      items2:[],
      userId,
      currUser,
      day_btncolor:["gainsboro","gainsboro","gainsboro","gainsboro","gainsboro","gainsboro","gainsboro"],
      days_selected:[],
      disableButton
    }; 
  } 
 
  public async componentDidMount(): Promise<void>
  {  console.log("Rendering..........");
    // get all the items from a sharepoint list
    var reacthandler=this;
    var col=this.props.WorkLoc_Column;
    sp.web.lists.getByTitle(this.props.WorkLoc).items.select(col).get().then(function(data){
      for(var k in data){
        items1.push({key:data[k].col, text:data[k].col});
      }
      reacthandler.setState({items1});
      console.log(items1);
      return items1;
    });
    this.getCurrentUser();
  }
  
  public render(): React.ReactElement<IRtwpFormProps> {
    
    return (
      <Form>
        <Label><h3>{this.props.description}</h3></Label>
  <Form.Group as={Row} controlId="formHorizontalEmail">
    <Form.Label column sm={4}>
    Employee Name
    </Form.Label>
    <Col sm={8}>
    <PeoplePicker
    context={this.props.context1}
    personSelectionLimit={1}
    groupName={""} // Leave this blank in case you want to filter from all users
    showtooltip={false}
    isRequired={true}
    disabled={false}
    selectedItems={this._getPeoplePickerItems}
    showHiddenInUI={false}
    principalTypes={[PrincipalType.User]}
    resolveDelay={1000}/>
    </Col>
  </Form.Group>

  <Form.Group as={Row} controlId="formHorizontalPassword">
    <Form.Label column sm={4}>
    Employee Id
    </Form.Label>
    <Col sm={8}>
      <Form.Control type="text" placeholder="Employee Id" readOnly={true} value={this.state.userId} />
    </Col>
  </Form.Group>
  <Form.Group as={Row} controlId="formHorizontalPassword">
    <Form.Label column sm={4}>
      Supervisor
    </Form.Label>
    <Col sm={8}>
      <Form.Control type="text" placeholder="SuperVisor Name" />
    </Col>
  </Form.Group>
  <Form.Group as={Row} controlId="formHorizontalRepEmp">
    <Form.Label column sm={4}>
      Represented Employee
    </Form.Label>
    <Col sm={8}>
  <Toggle  defaultChecked onText="On" offText="Off" onChange={_onChange} />
  </Col>
  </Form.Group>
  <Form.Group as={Row} controlId="formHorizontalWorkLoc">
    <Form.Label column sm={4}>
      Primary Work Location
    </Form.Label>
    <Col sm={8}>
  <Stack tokens={stackTokens}>
      <Dropdown placeholder="Select an option" onChange={this.onChange_Work} options={this.state.items1} />
    </Stack>
  </Col>
  </Form.Group>
  <Form.Group as={Row} controlId="formHorizontalBuild">
    <Form.Label column sm={4}>
      Building/Floor
    </Form.Label>
    <Col sm={8}>
    <Stack tokens={stackTokens}>
      <Dropdown placeholder="Select an option" options={this.state.items2} onChange={this.onChange_Building}/>
    </Stack>
  </Col>
  </Form.Group>
  <Form.Group as={Row} controlId="formHorizontalReason">
    <Form.Label column sm={4}>
      Reason for return
    </Form.Label>
    <Col sm={8}>
    <Stack tokens={stackTokens}>
      <Dropdown placeholder="Select an option" options={Reason} onChange={this.onChange_Reason}/>
    </Stack>
  </Col>
  </Form.Group>
  <Form.Group as={Row} controlId="formHorizontalAddInfo">
    <Form.Label column sm={4}>
    Provide additional information on why you want to return
    </Form.Label>
    <Col sm={8}>
  <TextField multiline resizable={false} onChange={this.getAddInfo}/>
  </Col>
  </Form.Group>
  <Form.Group as={Row} controlId="formHorizontalAddInfo">
    <Form.Label column sm={4}>
    Select Days
    </Form.Label>
    <Col sm={1}>
    <DefaultButton text="Mon" onClick={()=>this._dayClicked(0)} style={{backgroundColor:this.state.day_btncolor[0]}}/>
  </Col>
  <Col sm={1}>
    <DefaultButton text="Tue" onClick={()=>this._dayClicked(1)} style={{backgroundColor:this.state.day_btncolor[1]}}/>
  </Col> 
  <Col sm={1}>
    <DefaultButton text="Wed" onClick={()=>this._dayClicked(2)} style={{backgroundColor:this.state.day_btncolor[2]}}/>
  </Col> 
  <Col sm={1}>
    <DefaultButton text="Thur" onClick={()=>this._dayClicked(3)} style={{backgroundColor:this.state.day_btncolor[3]}}/>
  </Col>
  <Col sm={1}>
    <DefaultButton text="Fri" onClick={()=>this._dayClicked(4)} style={{backgroundColor:this.state.day_btncolor[4]}}/>
  </Col> 
  <Col sm={1}>
    <DefaultButton text="Sat" onClick={()=>this._dayClicked(5)} style={{backgroundColor:this.state.day_btncolor[5]}}/>
  </Col>
  <Col sm={1}>
    <DefaultButton text="Sun" onClick={()=>this._dayClicked(6)} style={{backgroundColor:this.state.day_btncolor[6]}}/>
  </Col>
  </Form.Group>
  <Form.Group as={Row} controlId="formHorizontalSubmitted">
    <Form.Label column sm={4}>
      Submitted by
    </Form.Label>
    <Col sm={8}>
    <Label>{this.state.currUser}</Label>
  </Col>
  </Form.Group>
  <Form.Group as={Row} controlId="formHorizontalAddInfo">
    <Form.Label column sm={4}>
    
    </Form.Label>
    <Col sm={8}>
    <Checkbox label="I agree to the Covid Compliance" onChange={this._onCovidComp} />
  </Col>
  </Form.Group> 
  <Form.Group as={Row} controlId="formHorizontalAddInfo">
    <Form.Label column sm={4}>
    
    </Form.Label>
    <Col sm={8}>
    <PrimaryButton text="Submit" onClick={this._submitClicked} disabled={this.state.disableButton}/>
  </Col>
  </Form.Group> 
</Form>
    );

    
  }
  private _getPeoplePickerItems=(items: any[])=> {
    var reacthandler=this;
    empName=items[0].text;
    sp.web.ensureUser(items[0].secondaryText).then((data)=>{
      console.log(data);
      userId=data.data.Id;
      reacthandler.setState({userId});
 });
     }
    
  private onChange_Work = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    items2.splice(0);
    var index = items1.indexOf(item);
    selectedWork=index+1;
    var reacthandler=this;
    sp.web.lists.getByTitle("FloorData").items.select('Title','Building_x002f_FloorNo','ID').get().then((data)=>{
      for(var k in data){
         if(item.key==data[k].Title)
        items2.push({key:data[k].ID, text:data[k].Building_x002f_FloorNo});
      }
      reacthandler.setState({items2});
      console.log(items2);
      return items2;
    });
  }
 private onChange_Building= (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
  var index = Number(item.key);
  selectedBuilding=index; 
 }
 private onChange_Reason=(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
  reason_Selected=item.text;
 }
  private getAddInfo= (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void=>{
   AddInfo=newText;

  };  
  private getCurrentUser=()=>{
    var reacthandler=this;
    sp.web.currentUser.get().then((data)=>{
      currUser=data.Title;
      reacthandler.setState({currUser});
    });
  }
  
  private _dayClicked=((day:number)=>{
    if(day_btncolor[day]=="gainsboro"){
      day_btncolor[day]="darkgray";
      switch (day) {
        case 0: days_selected.push("Monday"); break;
        case 1: days_selected.push("Tuesday"); break;
        case 2: days_selected.push("Wednesday"); break;
        case 3: days_selected.push("Thursday"); break;
        case 4: days_selected.push("Friday"); break;
        case 5: days_selected.push("Saturday"); break; 
        case 6: days_selected.push("Sunday"); break; 

        default:
          break;
      }
      this.setState({day_btncolor});
      this.setState({days_selected});
      console.log(days_selected);
    }
    else{
      day_btncolor[day]="gainsboro";
      this.setState({day_btncolor});
      switch (day) {
        case 0: var index = days_selected.indexOf("Monday");
                if (index !== -1) days_selected.splice(index, 1); break;
        case 1: var index = days_selected.indexOf("Tuesday");
                if (index !== -1) days_selected.splice(index, 1); break;
        case 2:var index = days_selected.indexOf("Wednesday");
                if (index !== -1) days_selected.splice(index, 1);break;
        case 3: var index = days_selected.indexOf("Thursday");
                if (index !== -1) days_selected.splice(index, 1); break;
        case 4: var index = days_selected.indexOf("Friday");
                if (index !== -1) days_selected.splice(index, 1); break;
        case 5: var index = days_selected.indexOf("Saturday");
                  if (index !== -1) days_selected.splice(index, 1); break; 
        case 6: var index = days_selected.indexOf("Sunday");
                  if (index !== -1) days_selected.splice(index, 1); break; 

        default:
          break;
      }
      
      this.setState({days_selected});
      console.log(days_selected);
    }
  });
  private _onCovidComp=((ev: React.FormEvent<HTMLElement>, isChecked: boolean)=>{
    if(isChecked){
   if((empName!=null)&&(selectedWork!=null)&&(selectedBuilding!=null)
      &&(AddInfo!=null)&&(days_selected!=null)&&(currUser!=null)&&(reason_Selected!=null)){
        disableButton=false;
        this.setState({disableButton});         
      }
     }
 });

  private _submitClicked(): void {
    sp.web.lists.getByTitle('Requests').items.add({
      Employee_x0020_Name: empName,
      Work_x0020_LocationId: selectedWork,
      Building_FloorId:selectedBuilding,
      Reason_x0020_for_x0020_Return:reason_Selected,
      Additional_x0020_Info:AddInfo,
      Days_Selected:days_selected.toString(),
      SubmittedBy:currUser
    
  });
  }
}
function _onChange(ev: React.MouseEvent<HTMLElement>, checked: boolean) {
  console.log('toggle is ' + (checked ? 'checked' : 'not checked'));
}


