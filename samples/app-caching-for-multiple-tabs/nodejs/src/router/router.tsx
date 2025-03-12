import * as React from "react";
import {
    BrowserRouter,
    Route,
    Routes,
    useNavigate
} from 'react-router-dom';
import * as microsoftTeams from "@microsoft/teams-js";
import Configure from "../components/configure";
import AppCacheTab from "../components/app-cache-tab";
import Index from "../components/index";
import Yellow from "../components/yellow";

const AppRoute = () => {
    const [appInitialized, setAppInitialized] = React.useState(false);

    const navigate = useNavigate();

    React.useEffect(() => {
        // Initialize the Microsoft Teams SDK
        const app = microsoftTeams.app;

        app.initialize().then(() => {
            microsoftTeams.teamsCore.registerBeforeUnloadHandler((readyToUnload: any) => {
                readyToUnload();
                console.log("sending readyToUnload to TEAMS");
                return true;
            });

            microsoftTeams.teamsCore.registerOnLoadHandler((data: any) => {
                 if (data.entityId) {
                    console.log("Load handler sending new entityId to TEAMS " + data.entityId);
                    if (data.entityId === 'yellow') {
                        navigate('/yellow');
                    } else if (data.entityId === 'red') {
                        navigate('/appCacheTab');
                    }
                }
            });

            setAppInitialized(true);
        }).catch(function (error: any) {
            console.error(error, "Could not initialize TeamsJS SDK.");
        });

        return () => {
            console.log("useEffect cleanup - Tab");
        };
    }, [navigate]);

    return (
        <React.Fragment>
            {appInitialized ? (
                <Routes>
                    <Route path="/" element={<Index />} />
                    <Route path="/configure" element={<Configure />}/>
                    <Route path="/appCacheTab" element={<AppCacheTab/>}/>
                    <Route path="/yellow" element={<Yellow />}/>
                </Routes>
            ) : null}
        </React.Fragment>
    );
};

export const App = () => (
    <BrowserRouter>
        <AppRoute />
    </BrowserRouter>
);