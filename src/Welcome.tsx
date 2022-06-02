import { Button, Container } from "react-bootstrap";
import { RouteComponentProps } from "react-router-dom";
import {
  AuthenticatedTemplate,
  UnauthenticatedTemplate,
  useMsal,
} from "@azure/msal-react";
import { useAppContext } from "./AppContext";
import { useEffect, useState } from "react";

export default function Welcome(props: RouteComponentProps) {
  const app = useAppContext();
  const { instance, accounts } = useMsal();
  const [accessToken, setAccessToken] = useState<string>("");

  const getAccessToken = () => {
    const request = {
      scopes: ["User.Read"],
      account: accounts[0],
    };

    instance
      .acquireTokenSilent(request as any)
      .then((response) => {
        setAccessToken(response.accessToken);
      })
      .catch((e) => {
        instance.acquireTokenPopup(request as any).then((response) => {
          setAccessToken(response.accessToken);
        });
      });
  };

  const callMsGraphWithAccessToken = async () => {
    const headers = new Headers();
    const accessToken = await app.authProvider?.getAccessToken();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);

    const options = {
      method: "GET",
      headers: headers,
    };

    const graphConfig = {
      graphMeEndpoint:
        "https://graph.microsoft.com/v1.0/me/drive/root/children",
    };

    return fetch(graphConfig.graphMeEndpoint, options)
      .then((response) => response.json())
      .catch((error) => console.log(error));
  };

  return (
    <div className="p-5 mb-4 bg-light rounded-3">
      <Container fluid>
        <h1>M365 App</h1>
        <p className="lead">
          This sample app shows how to use the Microsoft Graph API to access a
          user's data from React
        </p>
        <AuthenticatedTemplate>
          <div>
            <h4>Welcome {app.user?.displayName || ""}!</h4>
            <p>Use the navigation bar at the top of the page to get started.</p>
            <Button color="primary" onClick={callMsGraphWithAccessToken}>
              Get My Drive Items
            </Button>
          </div>
        </AuthenticatedTemplate>
        <UnauthenticatedTemplate>
          <Button color="primary" onClick={app.signIn}>
            Click here to sign in
          </Button>
        </UnauthenticatedTemplate>
      </Container>
    </div>
  );
}
