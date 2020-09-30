import React, { useEffect, useState } from 'react';
import logo from './logo.svg';
import './App.css';
import * as msTeams from '@microsoft/teams-js';
import { TeamsSSOHome } from './components/TeamsSSOHome';
import { TeamsMSALHome } from './components/TeamsMSALHome';
import { SharePointHome } from './components/SharePointHome';
import { WebAppHome } from './components/WebAppHome';

function App() {
  const [ isInitialized, setIsInitialized ] = useState<boolean>(false);
  const [ teamsContext, setTeamsContext ] = useState<msTeams.Context>();
  const [ spContext, setSpContext ] = useState<any>();

  const receiveSpContext = (evt: MessageEvent) => {
    if (!evt.origin.endsWith('.sharepoint.com')) return;
    setSpContext(evt.data);
  }

  const renderHomeScreen = () => {
    if (!isInitialized) return null;
    else if (teamsContext) return <TeamsSSOHome teamsContext={teamsContext} />
    // else if (teamsContext) return <TeamsMSALHome teamsContext={teamsContext} />
    else if (spContext) return <SharePointHome spContext={spContext} />
    else return <WebAppHome />
  }

  const refreshTeamsContext = () => msTeams.getContext((ctx) => setTeamsContext(ctx));

  useEffect(() => {
    window.addEventListener('message', receiveSpContext, false);

    msTeams.initialize(() => {
      refreshTeamsContext();
      msTeams.registerOnThemeChangeHandler(refreshTeamsContext);
    });
    
    setTimeout(() => setIsInitialized(true), 10);
    return () => window.removeEventListener('message', receiveSpContext);
  }, []);

  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
      </header>
      {renderHomeScreen()}
    </div>
  );
}

export default App;
