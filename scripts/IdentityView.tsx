import "../css/IdentityView.scss";

import * as React from "react";
import { parseUniquefiedIdentityName } from "./Helpers";

import { IdentityReference } from "TFS/WorkItemTracking/Contracts";

export interface IIdentityViewProps {
    identityDistinctName?: string;
    identity?: IdentityReference
}

export var IdentityView: React.StatelessComponent<IIdentityViewProps> =
    (props: IIdentityViewProps): JSX.Element => {
        let displayName = "";
        let imageUrl = "";

        if (props.identityDistinctName) {
            let identityRef = parseUniquefiedIdentityName(props.identityDistinctName);
            displayName = identityRef.displayName;
            imageUrl = identityRef.imageUrl;
        }
        else if (props.identity) {
            displayName = props.identity.name;
            imageUrl = `${VSS.getWebContext().host.uri}/_api/_common/IdentityImage?id=${props.identity.id}`;
        }
        
        if (!displayName) {
            return <div className="overflow-ellipsis" title="" />;
        }
        else {
            return (
                <div className="identity-view overflow-ellipsis" title={props.identityDistinctName}>
                    {imageUrl !== "" && (<img src={imageUrl} />)}
                    <span>{displayName}</span>
                </div>
            );
        }
}