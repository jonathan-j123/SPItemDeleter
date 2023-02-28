import * as React from 'react';
import styles from './ProductivityApp.module.scss';
import { IProductivityAppProps } from './IProductivityAppProps';

// import interfaces
import { IResponseItem } from "./interfaces";

import { getSP } from "../pnpjsConfig";
import { SPFI, spfi } from "@pnp/sp";
import { Logger, LogLevel } from "@pnp/logging";
import { useState } from 'react';

export interface IIPnPjsExampleState {
  items: IResponseItem[];
  errors: string[];
  disabled: boolean;
  submitDisabled: boolean;
}

const InputList = (props: { getItems: (arg0: string) => void, disabled: boolean, deleteItems: (listName:string) => void, submitDisabled:boolean }):JSX.Element => {
  const [listTitle, setListTitle] = useState("")
  return (
    <div>
      <label htmlFor="listTitle">List Name:</label>
      <input type="text" name='listTitle' value={listTitle} onChange={(e)=> setListTitle(e.target.value)} />
      <button disabled={props.submitDisabled} onClick={()=>props.getItems(listTitle)}>Submit</button>
      <p>Status: {props.disabled?<span style={{color:'red'}}>Not Connected</span>:<span style={{color:'green'}}>Connected</span>}</p>
      <button disabled={props.disabled} onClick={()=>props.deleteItems(listTitle)}>Remove all items</button>
      <br/>
      <p id="result" />
    </div>
  )
}

export default class PnPjsExample extends React.Component<IProductivityAppProps, IIPnPjsExampleState> {
  private LOG_SOURCE = "🅿PnPjsExample";
  private _sp: SPFI;

  constructor(props: IProductivityAppProps) {
    super(props);
    // set initial state
    this.state = {
      items: [],
      errors: [],
      disabled: true,
      submitDisabled: false,
    };
    this._sp = getSP();
  }

  public componentDidMount(): void {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    // this.getItems();
  }

  public render(): React.ReactElement<IProductivityAppProps> {

    return (
      <div className={styles.productivityApp}>
        <h1>LIST ITEMS REMOVER</h1>
        <InputList getItems={this.getItems} disabled={this.state.disabled} deleteItems={this.deleteItems} submitDisabled={this.state.submitDisabled}/>
      </div >
    );
  }

  

  private getItems = async (listName:string): Promise<void> => {
    document.getElementById("result").innerHTML = ""
    try {
      const sp = spfi(this._sp);
  
      const response: IResponseItem[] =  await sp.web.lists
      .getByTitle(listName)
      .items
      .select("Id").top(1000)();
      const items: IResponseItem[] = response.map((item: IResponseItem) => {
        return {
          Id: item.Id
        };
      });
  
      // Add the items to the state
      this.setState({ items });
      console.log(this.state.items.length)
  this.setState({disabled: false});
    } catch (err) {
      Logger.write(`${this.LOG_SOURCE} (_readAllFilesSize) - ${JSON.stringify(err)} - `, LogLevel.Error);
      this.setState({disabled: true});
    }
  }

  private  deleteItems = async (listName:string):Promise<void> => {
    // let timer = 60000;
    try {
    await this.getItems(listName);
    this.setState({disabled: true});
    this.setState({submitDisabled: true});
    
    this.state.items.forEach( async item => {
      const sp = spfi(this._sp);
      await sp.web.lists.getByTitle(listName).items.getById(item.Id).delete();
      
    });
    
    
      
      //  // eslint-disable-next-line no-unused-expressions
      //  this.state.items[0].Id === undefined && this.deleteItems(listName);
      //  return setTimeout(() => {
      //   console.log(`${this.state.items.length} items have been deleted.`);
      //   // eslint-disable-next-line @typescript-eslint/no-floating-promises
      //   this.deleteItems(listName);
      //  }, timer);
    this.setState({submitDisabled: false});
    }
    catch(err) {
      document.getElementById("result").innerHTML = "<span><strong style='background: green; color: white; padding: 0 4px; border-radius: 2px'>Success!</strong> The items have been removed from the '" + listName + "' SharePoint list.</span>";
      this.setState({submitDisabled: false});
    }
  }
}

