import * as React from 'react';
import styles from './SharePointLists.module.scss';
import { ISharePointListsProps } from './ISharePointListsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ISharePointListsState } from './ISharePointListsState';

require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');

export default class SharePointLists extends React.Component<ISharePointListsProps, ISharePointListsState, {}> {
  constructor(props?:ISharePointListsProps, context?:any){
    super(props,context);
    this.state = {
      listTitles: [],
      loadingLists: false,
      error: null
    };
    this.getListsTitles = this.getListsTitles.bind(this);
  }

  private getListsTitles():void{
    this.setState({
      listTitles: [],
      loadingLists: false,
      error: null
    });

    const context: SP.ClientContext = new SP.ClientContext(this.props.siteURL);
    const lists: SP.ListCollection = context.get_web().get_lists();
    context.load(lists, 'Include(Title)');
    context.executeQueryAsync((sender: any, args: SP.ClientRequestSucceededEventArgs): void => {
      const listEnumerator: IEnumerator<SP.List> = lists.getEnumerator();
      const titles: string[] = [];
      while(listEnumerator.moveNext()){
        const list: SP.List = listEnumerator.get_current();
        titles.push(list.get_title());
      }

      this.setState((prevState:ISharePointListsState, props: ISharePointListsProps): ISharePointListsState =>{
        prevState.listTitles = titles;
        prevState.loadingLists = false;
        return prevState;
      });

    }
    ,(sender: any, args: SP.ClientRequestFailedEventArgs): void=> {
        this.setState({
          listTitles: [],
          loadingLists: false,
          error: args.get_message()
        });
    }
  );
}
  public render(): React.ReactElement<ISharePointListsProps> {
    const titles:JSX.Element[] = this.state.listTitles.map((listTitle:string, index:number,listTitles:string[]): JSX.Element => {
      return <li key={index}> {listTitle} </li>;
    });
    return (
      <div className={ styles.sharePointLists }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a className={ styles.button } onClick={this.getListsTitles} role="button">
                <span className={ styles.label }>Get Lists Titles</span>
              </a> <br />
              {this.state.loadingLists &&
              <span>Loading lists...</span>}
              {this.state.error &&
              <span>An error occured while loading lists: {this.state.error}</span>}
              {this.state.error === null && titles && 
              <ul>
                {titles}
                </ul>}
            </div>
          </div>
        </div>
      </div>
    );
  }
}
