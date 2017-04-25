import * as React from "react";

import {ContextPropTypes} from "./Interfaces/FluxComponent";

export class FluxComponent<TP, TS> extends React.Component<TP, TS> {
    static childContextTypes = ContextPropTypes;

    protected _context: IQueriesHubContext;

    constructor(props: TP, context?: any) {
        super(props, context);        

        this._context = this._ensureQueriesHubContext();
    }
}