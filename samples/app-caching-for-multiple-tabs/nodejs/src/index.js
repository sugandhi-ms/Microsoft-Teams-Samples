import { FluentProvider, teamsLightTheme } from '@fluentui/react-components';
import React from 'react';
import ReactDOM from 'react-dom';
import './index.css';
import { App } from './router/router';

ReactDOM.render(
    <FluentProvider theme={teamsLightTheme}>
        <React.StrictMode>
            <App />
        </React.StrictMode>
    </FluentProvider>, document.getElementById('root')
);
