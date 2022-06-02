import { BrowserRouter as Router, Route } from "react-router-dom";
import { Container } from "react-bootstrap";
import { MsalProvider } from "@azure/msal-react";
import { IPublicClientApplication } from "@azure/msal-browser";

import ProvideAppContext from "./AppContext";
import ErrorMessage from "./ErrorMessage";
import NavBar from "./NavBar";
import Welcome from "./Welcome";
import "bootstrap/dist/css/bootstrap.css";
import Calendar from "./Calendar";
import NewEvent from "./NewEvent";
import Documents from "./Documents";
import DocumentDetail from "./DocumentDetail";
import Sites from "./Sites";
import Lists from "./Lists";

type AppProps = {
  pca: IPublicClientApplication;
};

export default function App({ pca }: AppProps) {
  return (
    <MsalProvider instance={pca}>
      <ProvideAppContext>
        <Router>
          <div>
            <NavBar />
            <Container>
              <ErrorMessage />
              <Route
                exact
                path="/"
                render={(props) => <Welcome {...props} />}
              />
              <Route
                exact
                path="/calendar"
                render={(props) => <Calendar {...props} />}
              />
              <Route
                exact
                path="/newevent"
                render={(props) => <NewEvent {...props} />}
              />
              <Route
                exact
                path="/sites/:siteName/lists/:listId/docs"
                render={(props) => <Documents {...props} />}
              />
              <Route
                exact
                path="/sites/:siteName/lists/:listId/docs/:driveItemId/:listItemId"
                render={(props) => <DocumentDetail {...props} />}
              />
              <Route
                exact
                path="/sites"
                render={(props) => <Sites {...props} />}
              />
              <Route
                exact
                path="/sites/:siteName/lists/"
                render={(props) => <Lists {...props} />}
              />
            </Container>
          </div>
        </Router>
      </ProvideAppContext>
    </MsalProvider>
  );
}
