import * as React from 'react';
import { useEffect, useState } from 'react';
import { AuthResponse } from 'msal';
import auth from '../auth';

import { Stack } from '@fluentui/react/lib/Stack';

export interface ISharePointHomeProps {
  spContext: any
}

export const SharePointHome: React.FC<ISharePointHomeProps> = ({ spContext }) => {
  const [ authResponse, setAuthResponse ] = useState<AuthResponse>();
  const [ isAuthenticated, setIsAuthenticated ] = useState<boolean>(false);

  useEffect(() => {
    try {
      console.log(`SharePoint Context`, spContext);
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

  const onClickSignIn = async (): Promise<void> => {
    return auth.attemptPopupSignIn().then((resp) => {
      if (resp) {
        setIsAuthenticated(true);
        setAuthResponse(resp);
      }
    });
  }

  return (
    <section>
      <h2>SharePoint</h2>
      {isAuthenticated && authResponse && (
        <h3>Hello, {authResponse.idToken.name || authResponse.idToken.preferredName}</h3>
      )}
      {spContext && (
        <Stack>
          <h2>SharePoint Context</h2>
          <Stack tokens={{ childrenGap: 10 }}>
            <Stack>
              <strong>Web URL</strong>
              <span>{spContext.web.absoluteUrl}</span>
            </Stack>
            <Stack>
              <strong>Web ID</strong>
              <span>{spContext.web.id._guid}</span>
            </Stack>
          </Stack>
        </Stack>
      )}
      {!isAuthenticated && (
        <button onClick={onClickSignIn}>Sign In</button>
      )}
    </section>
  )
};