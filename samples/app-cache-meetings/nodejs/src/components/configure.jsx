// <copyright file="configure.jsx" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

import { useEffect, useState } from "react";
import * as microsoftTeams from "@microsoft/teams-js";

// Configure page.
const Configure = props => {
    const [selectedComponent, setSelectedComponent] = useState("Component 1");

    useEffect(() => {
        microsoftTeams.app.initialize().then(() => {
            microsoftTeams.app.notifySuccess();

            microsoftTeams.pages.config.registerOnSaveHandler(function (saveEvent) {
                const contentUrl = selectedComponent === "Component1" 
                ? `${window.location.origin}/appCacheTab` 
                : `${window.location.origin}/second`;

                microsoftTeams.pages.config.setConfig({
                    //entityId : Generating a random id so that each tab instance has a unique ID.
                    entityId: "AppInstance_" + selectedComponent,
                    contentUrl,
                    suggestedTabName: selectedComponent,
                    websiteUrl: contentUrl,
                });
                saveEvent.notifySuccess();
            });
            microsoftTeams.pages.config.setValidityState(true);
        });
    }, [selectedComponent]);
    return (
        <div>
            <h2>App Caching</h2>
            <h3>This sample app only supports app caching.</h3>
            <p>Please click save button to proceed.</p>
            <label htmlFor="componentSelect">Select Component:</label>
            <select 
                id="componentSelect" 
                value={selectedComponent} 
                onChange={(e) => setSelectedComponent(e.target.value)}
            >
                <option value="Component1">Component 1</option>
                <option value="Component2">Component 2</option>
            </select>
        </div>
    );
};

export default Configure;