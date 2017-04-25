import * as React from "react";

export const ContextPropTypes: React.ValidationMap<any> = {
    actions: React.PropTypes.object.isRequired,
    stores: React.PropTypes.object.isRequired,
    actionsCreator: React.PropTypes.object.isRequired
};