import "../css/InfoLabel.scss";

import * as React from "react";
import {Label} from "OfficeFabric/Label";
import {Icon} from "OfficeFabric/Icon";

export interface IInfoLabelProps {
    label: string;
    info: string;
}

export var InfoLabel: React.StatelessComponent<IInfoLabelProps> =
    (props: IInfoLabelProps): JSX.Element => {        
        return (
            <div className="info-label">
                <Label className="info-label-text">{props.label}</Label>
                <span className="rich-tooltipped rich-tooltipped-se" aria-label={props.info}>
                    <Icon className="info-icon" iconName="Info" />
                </span>
            </div>
        );
}