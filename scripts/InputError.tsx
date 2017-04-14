import "../css/InputError.scss";

import * as React from "react";
import {Icon} from "OfficeFabric/Icon";

export interface IInputErrorProps {
    error: string;
}

export var InputError: React.StatelessComponent<IInputErrorProps> =
    (props: IInputErrorProps): JSX.Element => {        
        return (
            <div className="input-error">
                <Icon className="error-icon" iconName="Error" />
                <span>{props.error}</span>
            </div>
        );
}