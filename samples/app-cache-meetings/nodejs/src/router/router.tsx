// <copyright file="router.js" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>


import React, { useState } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import {
    BrowserRouter,
    Route,
    Routes,
    useNavigate
} from 'react-router-dom';
import Configure from "../components/configure";
import AppCacheTab from "../components/app-cache-tab";
import Index from "../components/index";
import Second from "../components/second";

const AppContent = () => {
    let app = microsoftTeams.app;
    const navigate = useNavigate();

    React.useEffect(() => {
        app.initialize().then(app.getContext).then((context: any) => {
            app.notifySuccess();

            // Get default theme from app context and set app-theme
            let defaultTheme = context.app.theme;

            switch (defaultTheme) {
                case 'dark':
                    console.log("theme-dark");
                    break;
                default:
                    console.log('theme-light');
            }

            // Handle app theme when 'Teams' theme changes
            microsoftTeams.app.registerOnThemeChangeHandler(function (theme) {
                switch (theme) {
                    case 'dark':
                        console.log('Register theme-dark');
                        break;
                    case 'default':
                        console.log('Register theme-light');
                        break;
                    case 'contrast':
                        console.log('Register theme-contrast');
                        break;
                    default:
                        return console.log('Register Default theme-dark');
                }
            });

                microsoftTeams.teamsCore.registerBeforeUnloadHandler((readyToUnload: any) => {
                   readyToUnload();

                    return true;
                });

                microsoftTeams.teamsCore.registerOnLoadHandler((data: any) => {
                    console.log(data.contentUrl, data.entityId);
                    if (data.entityId === 'AppInstance_Component1') {
                        navigate('/second');
                    } else {
                        navigate('/appCacheTab');
                    }
                    app.notifySuccess();
                });
        }).catch(function (error: any) {
            console.log(error, "Could not register handlers.");
        });

        return () => {
            console.log("useEffect cleanup - Tab");
        };

    }, []);

    return null;
};
export const AppRoute = () => {
    return (
        <React.Fragment>
            <BrowserRouter>
            <AppContent/>
                <Routes>
                    <Route path="/" element={<Index />} />
                    <Route path="/configure" element={<Configure />}/>
                    <Route path="/appCacheTab" element={<AppCacheTab />}/>
                    <Route path="/second" element={<Second />}/>
                </Routes>
            </BrowserRouter>
        </React.Fragment>
    );
};