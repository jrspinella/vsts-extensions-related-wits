import "../css/app.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { IWorkItemNotificationListener, IWorkItemChangedArgs, IWorkItemFieldChangedArgs, IWorkItemLoadedArgs } from "TFS/WorkItemTracking/ExtensionContracts";
import { WorkItemFormService, IWorkItemFormService } from "TFS/WorkItemTracking/Services";
import { WorkItem } from "TFS/WorkItemTracking/Contracts";

import {Loading} from "./Loading";
import {MessagePanel, MessageType} from "./MessagePanel";
import {InputError} from "./InputError";

interface IRelatedWitsState {
    loading: boolean;
    listItems: IListItem[];
    isWorkItemLoaded?: boolean;
    isNew?: boolean;
}

interface IListItem {
    workItem: WorkItem;
    isLinked: boolean;
}

export class RelatedWits extends React.Component<void, IRelatedWitsState> {
    constructor(props: void, context?: any) {
        super(props, context);

        VSS.register(VSS.getContribution().id, {
            onLoaded: (args: IWorkItemLoadedArgs) => {
                if (args.isNew) {
                    this.setState({...this.state, isWorkItemLoaded: true, isNew: true});
                }
                else {
                    this._refreshList();
                }
            },
            onUnloaded: (args: IWorkItemChangedArgs) => {
                this.setState({...this.state, isWorkItemLoaded: false, workItems: []});
            },
            onSaved: (args: IWorkItemChangedArgs) => {
                this._refreshList();
            },
            onRefreshed: (args: IWorkItemChangedArgs) => {
                this._refreshList();
            },
            onReset: (args: IWorkItemChangedArgs) => {
                this._refreshList();
            },
        } as IWorkItemNotificationListener);

        this.state = {
            loading: true,
            listItems: []
        } as IRelatedWitsState;
    }

    public render(): JSX.Element {
        if (!this.state.isWorkItemLoaded) {
            return null;
        }
        else if (this.state.loading) {
            return <Loading />;
        }
        else if (this.state.isNew) {
            return <MessagePanel message="Please save the workitem to get the list of related work items." messageType={MessageType.Info} />;
        }
        else if (!this.state.listItems || this.state.listItems.length === 0) {
            return <MessagePanel message="No related work items found." messageType={MessageType.Info} />;
        }
        else {
            
        }
    }

    private async _refreshList(): Promise<void> {
        this.setState({...this.state, isWorkItemLoaded: true, isNew: false, loading: true});


        this.setState({...this.state, loading: false, listItems: []});
    }
}

export function init() {
    ReactDOM.render(<RelatedWits />, $("#ext-container")[0]);
}