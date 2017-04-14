import "../css/MessagePanel.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import {Icon, IconName} from "OfficeFabric/Icon";
import {Label} from "OfficeFabric/Label";

interface IMessagePanelInputs {
    Message: string;
    MessageType: string;
}

interface IMessagePanelProps {
    message: string;
    messageType: MessageType;
}

export enum MessageType {
    Error,
    Warning,
    Info,
    Success
}

export var MessagePanel: React.StatelessComponent<IMessagePanelProps> =
    (props: IMessagePanelProps): JSX.Element => {
        let className = "message-panel";
        let iconName: IconName;

        switch (props.messageType) {
            case MessageType.Error:
                iconName = "StatusErrorFull";
                className += " message-error";
                break;
            case MessageType.Warning:
                iconName = "Warning";
                className += " message-warning";
                break;
            case MessageType.Success:
                iconName = "SkypeCircleCheck";
                className += " message-success";
                break;
            default:
                iconName = "Info";
                className += " message-info";
                break;
        }

        return (
            <div className={className}>
                <Icon className="icon" iconName={iconName} />
                <Label className="message-text">{props.message}</Label>
            </div>
        );
}