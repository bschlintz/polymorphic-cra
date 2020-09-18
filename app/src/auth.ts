import {
  AuthResponse, 
  Configuration, 
  errorReceivedCallback, 
  tokenReceivedCallback, 
  UserAgentApplication
} from "msal";

const msalConfig: Configuration = {
  auth: {
      clientId: process.env.REACT_APP_CLIENTID!,
      authority: process.env.REACT_APP_AUTHORITY!
  },
  cache: {
      cacheLocation: "localStorage"
  }
};

const scopes = ['user.read'];

export const attemptSilentSignIn = async (): Promise<AuthResponse | null> => {
  let authResponse = null;
  try {
    if (msalApp.getAccount()) {
      authResponse = await msalApp.acquireTokenSilent({ scopes });
    }
  }
  catch (error) {
    console.log(`error during silent sign in`, error);
  }
  finally {
    return authResponse;
  }
};

export const attemptRedirectSignIn = () => {
  msalApp.loginRedirect({ scopes });
}

export const attemptPopupSignIn = async (): Promise<AuthResponse | null> => {
  let authResponse = null;
  try {
    authResponse = await msalApp.loginPopup({ scopes });
  }
  catch (error) {
    console.log(`error during popup sign in`, error);
  }
  finally {
    return authResponse;
  }
}

export const signOut = () => {
  msalApp.logout();
}

const handleTokenReceived: tokenReceivedCallback = (response) => {
  console.log(`authentication successful`, response);
}

const handleErrorReceived: errorReceivedCallback = (authError, accountState) => {
  console.log(`authentication failed`, authError, accountState);
}

const msalApp = new UserAgentApplication(msalConfig);
msalApp.handleRedirectCallback(
  (response) => handleTokenReceived(response),
  (error, state) => handleErrorReceived(error, state)
);

export default {
  attemptSilentSignIn,
  attemptRedirectSignIn,
  attemptPopupSignIn,
  signOut,
  msalApp
}