import { Link, RouteComponentProps, useParams } from "react-router-dom";
import { useAppContext } from "./AppContext";
import { createList, getAllLists, searchSites } from "./GraphService";
import { useEffect, useState } from "react";
import "./Documents.css";
import { Button, Card, Col, Form, Row } from "react-bootstrap";

export default function Lists(props: RouteComponentProps) {
  const app = useAppContext();
  const { siteName } = useParams<{ siteName: string }>();
  const [lists, setLists] = useState<Array<any>>([]);
  const [siteId, setSiteId] = useState<string>();
  const [name, setName] = useState<string>("");

  useEffect(() => {
    const getSiteDetails = async () => {
      const siteDetails = await searchSites(app.authProvider!, siteName);
      setSiteId(siteDetails[0].id);
    };
    getSiteDetails();
  }, [app.authProvider, siteName]);

  const getLists = async () => {
    const lists = await getAllLists(app.authProvider!, siteId as string);
    setLists(lists);
  };

  useEffect(() => {
    if (siteId) getLists();
  }, [siteId]);

  const createNewList = async () => {
    const list = await createList(app.authProvider!, siteId as string, {
      displayName: name,
      list: {
        template: "documentLibrary",
      },
    });
    if (list?.id) {
      getLists();
    }
  };

  return (
    <div>
      <Row xs={1} md={2} className="g-4 align-items-center">
        <Col>
          <Form.Group className="mb-3">
            <Form.Label>List Name</Form.Label>
            <Form.Control
              value={name}
              onChange={(e: any) => setName(e.target.value)}
            />
          </Form.Group>
        </Col>
        <Col>
          <Button disabled={name === ""} onClick={createNewList}>
            Create New Document List
          </Button>
        </Col>
      </Row>
      <Row xs={1} md={4} sm={2} className="g-4">
        {lists.map((list: any) => (
          <Col>
            <Card>
              <Card.Header>{list.displayName}</Card.Header>
              <Card.Body>
                <Card.Text>{list.description}</Card.Text>
                <Link to={`/sites/${siteName}/lists/${list.id}/docs`}>
                  Documents
                </Link>
              </Card.Body>
            </Card>
          </Col>
        ))}
      </Row>
    </div>
  );
}
