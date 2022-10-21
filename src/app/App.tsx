import React from 'react';
//import './App.css'; // ==> ../index.html

import DatasetAppBar from './components/DatasetAppBar';
import DatasetForm from './components/DatasetForm';

export default function App() {
    return (
        <div className="App">
            <DatasetAppBar />
            <DatasetForm />
        </div>
    );
}
