import { AuthResponse } from 'msal';
import * as React from 'react';
import { useEffect, useState } from 'react';
import auth from '../auth';

export interface IWebAppHomeProps {
}

export const WebAppHome: React.FC<IWebAppHomeProps> = ({ }) => {
  const [ authResponse, setAuthResponse ] = useState<AuthResponse>();
  const [ isAuthenticated, setIsAuthenticated ] = useState<boolean>(false);

  useEffect(() => {
    try {
      auth.attemptSilentSignIn().then((resp) => {
        if (resp) {
          setIsAuthenticated(true);
          setAuthResponse(resp);
        }
      })
    }
    catch (error) {
      console.log(`Silent authentication failed`, error);
    }
  }, []);
  
  const onClickSignIn = () => {
    auth.attemptRedirectSignIn();
  }

  const onClickSignOut = () => {
    auth.signOut();
  }

  return (
    <section>
      <h2>Web App</h2>
      {isAuthenticated && authResponse && (
        <h3>Hello, {authResponse.idToken.name || authResponse.idToken.preferredName}</h3>
      )}
      {isAuthenticated 
        ? <button onClick={onClickSignOut}>Sign Out</button>
        : <button onClick={onClickSignIn}>Sign In</button>
      }
    </section>
  )
};