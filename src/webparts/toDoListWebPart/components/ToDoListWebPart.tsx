import * as React from 'react';
import styles from './ToDoListWebPart.module.scss';
import { IToDoListWebPartProps } from './IToDoListWebPartProps';
import { IToDoListWebPartState } from './IToDoListWebPartState';
import { escape } from '@microsoft/sp-lodash-subset';
import { divProperties, List, TagItemSuggestion } from 'office-ui-fabric-react';
import { cloneDeep } from '@microsoft/sp-lodash-subset';

import { initializeIcons } from '@fluentui/react/lib/Icons';
initializeIcons();
import 'office-ui-fabric-react/dist/css/fabric.css'

import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { useBoolean } from '@fluentui/react-hooks';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { TextField, ITextFieldStyles } from '@fluentui/react/lib/TextField';
import {
  DatePicker,
  DayOfWeek,
  Dropdown,
  IDropdownOption,
  mergeStyles,
  defaultDatePickerStrings,
  DropdownMenuItemType,
  IDropdownStyles,
} from '@fluentui/react';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import { Web } from '@pnp/sp/webs';
import "@pnp/sp/lists"



export default class ToDoListWebPart extends React.Component<IToDoListWebPartProps, IToDoListWebPartState> {
  constructor(props) {
    super(props);

    this.state = {
      items:[],
      showModal: false,
      activeItem: null,
      activeIndex: -1,
      showPanel: false,
      tempItem: {
        Title:'',
        Description:'',
        Status: 'Not Started',
        StartDate: new Date(),
      },
    };
  }

  public componentDidMount(): void {

  }

  public render(): React.ReactElement<IToDoListWebPartProps> {
    
    const d = new Date().toLocaleDateString();
    const {
      description,
      hasTeamsContext,
    } = this.props;
    const { items, showModal, activeItem, showPanel, tempItem } = this.state;

    return (
      <section className={`${styles.toDoListWebPart} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={"ms-Grid"}>

          <div className={'ms-Grid-row'}>

            <div className={'ms-Grid-col ms-sm12'}>
              {/* <h1>Hello World</h1>
              <h3>Today is: {d}</h3> */}
              <span>To Do List:</span>
            </div>
            <br/><br/>

            <div className={'ms-Grid-col ms-sm12'}>
              <PrimaryButton
                className={styles.primaryButton}
                label='Add Item'
                text='Add Item'
                // onClick={() => {alert("Hello World");}}
                onClick={() => {
                  // const item = {
                  //   Title: 'Hello World',
                  //   Description: 'Testing',
                  //   Status: 'Not Started',
                  //   DueDate: new Date().toLocaleDateString()
                  // };

                  // items.push(item);

                  // this.setState({items}, () => {
                  //   console.log('updating state', this.state)});
                  this.setState({showPanel: true})
                }}
              />
              <br/><br/>
            </div>

            <div className='ms-Grid-col ms-sm12'>

                <List
                  items={cloneDeep(items)}
                  onRenderCell={(item?: any, index?: number, isScrolling?: boolean) =>{
                    return (
                      <div className={'ms-Grid-row' + styles.divColor} style={{marginBottom: '10px', border: '1px ridge black'}} >

                        <div className='ms-Grid ms-sm8'>
                          <div className='ms-Grid-row ms-sm4'>
                              ID: {index + 1}
                          </div>
                          <div className='ms-Grid-row ms-sm4'>
                              Name: {item.Title}
                          </div>
                          <div className='ms-Grid-row ms-sm4'>
                              Status: {item.Status}
                          </div>
                          <div className='ms-Grid-row ms-sm4'>
                              Start Date: {item.StartDate}
                          </div>
                        </div><br/>

                        <div className='ms-Grid ms-sm4'>
                          <div className='ms-Grid-row ms-sm12' style={{margin:'5px auto'}}>
                            <div className='ms-Grid-row ms-sm2'>
                              <DefaultButton
                                style={{background: '#00b7c3', width:'100%', paddingTop:'15px 10px'}}
                                iconProps={{iconName: 'View'}}
                                onClick={
                                  () => {
                                    this.setState({
                                      showModal: true,
                                      activeItem: item,
                                      activeIndex: index
                                    });
                                  }
                                }
                              />
                            </div>
                          </div>

                          <div className='ms-Grid-row ms-sm12' style={{margin:'5px auto'}}>
                            <div className='ms-Grid-row ms-sm2'>
                              <DefaultButton
                                  style={{background: '#d83b01',width:'100%', paddingTop:'15px 10px'}}
                                  iconProps={{iconName: 'Delete'}}
                                  onClick={() => {
                                      const res = item.filter((it: any, num: number) => {
                                        if (index !== num) {
                                          return it;
                                        }
                                      });
                                      this.setState({items: cloneDeep(res)});
                                  }}
                              />
                            </div>
                          </div>

                        </div>

                      </div>
                    );
                  }}
                />

                <br/><br/>
            </div>

          </div>

{/*! Modal Section */}
          <Dialog
            hidden={!showModal}
            modalProps={{isBlocking: false,
             styles: { main: { maxWidth: 450 }
            }
            }}
            onDismiss={() => this.setState({showModal: false, activeItem:null, activeIndex: -1})}
            dialogContentProps={
              {
                type: DialogType.normal,
                title: 'Task Details'
              }
            }>

            <div className='ms-Grid-col ms-sm12'>
              <span style={{textAlign:'center'}}>
                {activeItem &&(
                  <div>

                    <b>Description:</b> {activeItem.Description}

                  </div>
                )}
              </span> 
            </div>

          </Dialog>

{/* Panel Section */}
          <Panel
            headerText='New Item'
            isOpen={showPanel}
            onDismiss={() => this.setState({showPanel: false})}
            // type={PanelType.medium}

          >
             <div className={'ms-Grid-col sm12'} style={{paddingLeft: '0px'}}>
              {/* name, description, status, due date */}
                <div className='ms-Grid-col ms-sm12' style={{paddingLeft: '0px'}}>
                  <TextField
                    placeholder='Enter Value here'
                    label='Title'
                    value={tempItem.Title}
                    onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newVal?: string) => {
                        tempItem.Title = newVal;

                        this.setState({tempItem});
                    }}
                  />
                </div>
                <br/><br/>

                <div className='ms-Grid-col ms-sm12' style={{paddingLeft: '0px'}}>
                  <TextField
                      placeholder='Enter Value here'
                      label='Description'
                      value={tempItem.Description}
                      onChange={(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newVal?: string) => {
                          tempItem.Description = newVal;

                          this.setState({tempItem});
                      }}
                      multiline
                      rows={6}
                  />
                </div>
                <br/><br/>

                <div className='ms-Grid-col ms-sm12' style={{paddingLeft: '0px'}}>
                  <Dropdown
                  placeholder='Select an option'
                  label='Status' 
                  options={[
                      {key: 'Not Started', text: 'Not Started'},
                      {key: 'In Progress', text: 'In Progress'},
                      {key: 'On Hold', text: 'On Hold'},
                      {key: 'Completed', text: 'Completed'},
                    ]}
                    
                    selectedKey={tempItem.Status}
                    onChange={(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption<any>, index?: number) => {
                      tempItem.Status = option.key;

                      this.setState({tempItem});
                    }}
                  />
                </div>

                <div className='ms-Grid-col ms-sm12' style={{paddingLeft: '0px'}}>
                  <DatePicker
                    label='Start Date'
                    placeholder="Select a date..."
                    value={tempItem.StartDate}
                    onSelectDate={(date: Date) => {
                      tempItem.StartDate = date;

                      this.setState({tempItem});
                    }}
                  />
                </div>
                


                <div className='ms-Grid-col ms-sm12' style={{paddingLeft: '0px', marginTop: '20px'}}>
                  
                  <div className='ms-Grid-col ms-sm3' style={{paddingLeft: '0px'}}>
                    <PrimaryButton
                      style={{width: '100%'}}
                      text='Save'
                      onClick={async() => {

                        //! save to sp list
                          //?for cross-site references only
                        const spWeb = Web('https://xqjdn.sharepoint.com/sites/DeveloperSite');

                        await spWeb.lists.getById('f35970d0-2358-4d47-86cb-3a02eb66a2ae').reserveListItemId
                        
                        //! query updates
                        //! refresh DOM
                        
                        //? items.push(tempItem);

                        //? this.setState({items,
                        //?  showPanel:false,
                        //? tempItem:{
                        //?     Title: '',
                        //?     Description: '',
                        //?     Status:'Not Started',
                        //?     StartDate: new Date(),
                        //?  }
                        //? });

                      }}
                    />
                  </div>

                  <div className='ms-Grid-col ms-sm3' style={{paddingLeft: '27px'}}>
                    <DefaultButton
                        style={{width: '100%'}}
                        text='Cancel'
                        onClick={() => {
                          this.setState({
                            showPanel:false,
                            tempItem:{
                              Title:'',
                              Description:'',
                              Status: 'Not Started',
                              StartDate: new Date(),
                            }
                          });
                        }}
                      />
                  </div>

                </div>

             </div>
          </Panel>

        </div> 
      </section>
    );
  }
}
