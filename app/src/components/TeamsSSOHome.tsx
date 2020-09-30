import * as React from 'react';
import * as msTeams from '@microsoft/teams-js';
import { useEffect, useState, useMemo } from 'react';
import decodeJwt from 'jwt-decode';

import { 
  Provider, teamsTheme, teamsDarkTheme, teamsHighContrastTheme,
  Table, TableCell, TableRow, Box, mergeThemes, ThemeInput 
} from '@fluentui/react-northstar';

export interface ITeamsHomeProps {
  teamsContext: msTeams.Context
}

const theme: ThemeInput = {
  componentVariables: {
    // 💡 `colorScheme` is the object containing all color tokens
    Box: (props: any) => ({
      // `brand` contains all design tokens for the `brand` color
      color: props.colorScheme.brand.foreground3,
      backgroundColor: props.colorScheme.brand.background3,
      // `foreground3` and `background3` are theme-dependent tokens that should
      // be used as value in styles, you can define own tokens 💪
    }),
  },
  componentStyles: {
    Box: {
      // 🚀 We recomend to use `colorScheme` from variables mapping
      root: (props: any) => ({
        color: props.variables.color,
        backgroundColor: props.variables.backgroundColor,
      }),
    },
  },
};

export const TeamsSSOHome: React.FC<ITeamsHomeProps> = ({ teamsContext }) => {
  const [ userToken, setUserToken ] = useState<any>();
  const [ isAuthenticated, setIsAuthenticated ] = useState<boolean>(false);

  const authSuccess = (token: string) => {
    const decodedToken = decodeJwt(token);
    console.log(`teams sso auth success`, decodedToken);
    setUserToken(decodedToken);
    setIsAuthenticated(true);
  };

  const authFail = (reason: string) => {
    console.log(`teams sso auth failed`, reason);
  };

  useEffect(() => {
    try {
      console.log(`Teams Context`, teamsContext);

      // To use Teams Tab SSO Auth
      msTeams.authentication.getAuthToken({
        successCallback: authSuccess,
        failureCallback: authFail
      });
    }
    catch (error) {
      console.log(`Silent authentication failed`, error);
    }
  }, []);

  const parentTheme = useMemo(() => {
    switch (teamsContext.theme)
    {
      case "dark": return teamsDarkTheme;
      case "teamsHighContrastTheme": return teamsHighContrastTheme;
      default: return teamsTheme;
    }
  }, [teamsContext.theme])

  return (
    <Provider theme={mergeThemes(parentTheme, theme)}>
      <Box content={teamsContext.theme}>
        <Box style={{ maxWidth: 1200, margin: "0 auto", padding: "20px"}}>
          <h2>Teams</h2>
          {isAuthenticated && userToken && (
            <h3>Hello SSO, {userToken.name} ({userToken.upn})</h3>
          )}
          {teamsContext && <>
            <h2>Teams Context</h2>
            <Table>
              <TableRow header>
                <TableCell><strong>Name</strong></TableCell>
                <TableCell><strong>Value</strong></TableCell>
              </TableRow>
              <TableRow>
                <TableCell>Theme</TableCell>
                <TableCell>{teamsContext.theme}</TableCell>
              </TableRow>
              <TableRow>
                <TableCell>Tab ID</TableCell>
                <TableCell>{teamsContext.entityId}</TableCell>
              </TableRow>
            </Table>
          </>}
        </Box>
      </Box>
    </Provider>
  )
};