
import 'babel-polyfill'
import React from 'react';
import ReactDOM from 'react-dom';

import HomePage from 'page/HomePage/HomePage.js';
import { observer } from 'mobx-react';

@observer
class App extends React.Component {
    render() {
        return (
            <div id="App_Conponent" className="Application">
                <HomePage></HomePage>
            </div>
        );
    }
}

ReactDOM.render(
    <App />,
    document.getElementById('app')
);
