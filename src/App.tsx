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
            </Container>
          </div>
        </Router>
      </ProvideAppContext>
    </MsalProvider>
  );
}
