import { Link, RouteComponentProps, useHistory } from "react-router-dom";
import { useAppContext } from "./AppContext";
import { AsyncTypeahead, Menu } from "react-bootstrap-typeahead";
import {
  createGroup,
  searchDocs,
  searchSites,
  uploadDocument,
} from "./GraphService";
import { useEffect, useState } from "react";
import "./Documents.css";
import { Alert, Button, Card, Col, Form, Row } from "react-bootstrap";

export default function Sites(props: RouteComponentProps) {
  const app = useAppContext();
  const [sites, setSites] = useState<Array<any>>([]);
  const [name, setName] = useState<string>("");
  const [nickName, setNickName] = useState<string>("");
  const [description, setDescription] = useState<string>("");
  const [show, setShow] = useState<boolean>(false);

  const getAllSites = async () => {
    const matchedSites = await searchSites(app.authProvider!, "*");
    setSites(matchedSites);
  };

  useEffect(() => {
    getAllSites();
  }, []);

  const createNewGroup = async () => {
    const group = await createGroup(
      app.authProvider!,
      name,
      description,
      nickName
    );
    if (group?.id) {
      getAllSites();
      setShow(true);
    }
  };

  return (
    <div>
      <Alert
        show={show}
        variant="success"
        onClose={() => setShow(false)}
        dismissible
      >
        New M365 Group successfully created!
      </Alert>
      <Row xs={1} md={4} className="g-4 align-items-center">
        <Col>
          <Form.Group className="mb-3">
            <Form.Label>Group Name</Form.Label>
            <Form.Control
              value={name}
              onChange={(e: any) => setName(e.target.value)}
            />
          </Form.Group>
        </Col>
        <Col>
          <Form.Group className="mb-3">
            <Form.Label>Description</Form.Label>
            <Form.Control
              value={description}
              onChange={(e: any) => setDescription(e.target.value)}
            />
          </Form.Group>
        </Col>
        <Col>
          <Form.Group className="mb-3">
            <Form.Label>Mail Nick Name</Form.Label>
            <Form.Control
              value={nickName}
              onChange={(e: any) => setNickName(e.target.value)}
            />
          </Form.Group>
        </Col>
        <Col>
          <Button
            disabled={name === "" || nickName === "" || description === ""}
            onClick={createNewGroup}
          >
            Create M365 Group
          </Button>
        </Col>
      </Row>
      <Row xs={1} md={4} sm={2} className="g-4">
        {sites.map((site: any) => (
          <Col>
            <Card>
              <Card.Header>{site.displayName}</Card.Header>
              <Card.Body>
                <Card.Text>{site.description}</Card.Text>
                <Link to={`/sites/${site.displayName}/lists`}>Site Lists</Link>
              </Card.Body>
            </Card>
          </Col>
        ))}
      </Row>
    </div>
  );
}
