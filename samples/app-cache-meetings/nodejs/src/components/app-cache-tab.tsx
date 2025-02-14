// <copyright file="app-cache-tab.tsx" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

import React from "react";
import "../components/index.css";
import { app } from "@microsoft/teams-js";

const AppCacheTab = () => {

    React.useEffect(()=>{
        app.notifySuccess();
    }, []);
    return (
        <div style={{backgroundColor: 'blue', width: '100px', height: '100px'}}>
            <h3>App Caching Sample</h3>
        </div>
    );
};

export default AppCacheTab;
