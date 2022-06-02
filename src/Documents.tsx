import { Link, RouteComponentProps, useParams } from "react-router-dom";
import { useAppContext } from "./AppContext";
import { AsyncTypeahead, Menu } from "react-bootstrap-typeahead";
import {
  createColumn,
  getAllColumns,
  searchDocs,
  searchDriveItem,
  searchSites,
  uploadDocument,
} from "./GraphService";
import { useEffect, useState } from "react";
import "./Documents.css";
import { Alert, Button, Col, Form, Row, Table } from "react-bootstrap";

export default function Documents(props: RouteComponentProps) {
  const { siteName, listId } = useParams<{
    siteName: string;
    listId: string;
  }>();
  const app = useAppContext();
  const [matchedDocs, setMatchedDocs] = useState<Array<any>>([]);
  const [show, setShow] = useState<boolean>(false);
  const [siteId, setSiteId] = useState<string>("");
  const [cols, setCols] = useState<any>();
  const [fileDetails, setFileDetails] = useState<{
    name: string;
    type: string;
    lastModifiedDate: Date;
  }>();
  const [name, setName] = useState<string>("");
  const [description, setDescription] = useState<string>("");

  useEffect(() => {
    const getSiteDetails = async () => {
      const siteDetails = await searchSites(app.authProvider!, siteName);
      setSiteId(siteDetails[0].id);
    };
    getSiteDetails();
  }, [app.authProvider, siteName]);

  const getAllCols = async () => {
    const columns = await getAllColumns(app.authProvider!, siteId, listId);
    setCols(columns);
  };

  useEffect(() => {
    if (siteId && listId) getAllCols();
  }, [siteId, listId]);

  const fileUpload = () => {
    const filename = fileDetails?.name;
    const filereader = new FileReader();
    filereader.onload = async (event) => {
      const resp = await uploadDocument(
        app.authProvider!,
        siteId,
        listId,
        filename as string,
        fileDetails
      );
      if (resp?.id) setShow(true);
    };
    filereader.readAsArrayBuffer(fileDetails as unknown as Blob);
  };

  const createNewColumn = async () => {
    const item = await createColumn(app.authProvider!, siteId, listId, {
      description: description,
      enforceUniqueValues: false,
      hidden: false,
      indexed: false,
      name: name,
      text: {
        allowMultipleLines: false,
        appendChangesToExistingText: false,
        linesForEditing: 0,
        maxLength: 255,
      },
    });
    if (item?.id) {
      getAllCols();
    }
  };

  return siteId !== "" ? (
    <div className="p-5 mb-4 bg-light rounded-3">
      <h3>{siteName}</h3>
      <Alert
        show={show}
        variant="success"
        onClose={() => setShow(false)}
        dismissible
      >
        Your upload was successful!
      </Alert>
      <AsyncTypeahead
        isLoading={false}
        placeholder="Search in the Sharepoint List"
        filterBy={() => true}
        onSearch={async (query) => {
          const docs = await searchDriveItem(app.authProvider!, query, listId);
          setMatchedDocs(docs);
        }}
        options={matchedDocs}
        renderMenu={(results, menuProps) => (
          <Menu {...menuProps}>
            {results.map((option: any) => (
              <div className="select-item">
                <Link
                  to={`/sites/${siteName}/lists/${listId}/docs/${option?.resource?.id}/${option?.resource?.parentReference?.sharepointIds?.listItemId}`}
                >
                  {option?.resource?.name}
                </Link>
              </div>
            ))}
          </Menu>
        )}
      />
      <div className="upload-container">
        <div>
          <Form.Group controlId="formFile" className="mb-3">
            <Form.Label>Upload File To The Sharepoint List</Form.Label>
            <Form.Control
              type="file"
              onChange={(e: any) => setFileDetails(e.target.files[0])}
            />
          </Form.Group>
          <Button
            disabled={fileDetails === undefined}
            className="button-format"
            onClick={() => fileUpload()}
          >
            Upload
          </Button>
        </div>
      </div>
      <Row xs={1} md={3} className="g-4 align-items-center">
        <Col>
          <Form.Group className="mb-3">
            <Form.Label>Column Name</Form.Label>
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
          <Button
            disabled={name === "" || description === ""}
            onClick={createNewColumn}
          >
            Create New Column
          </Button>
        </Col>
      </Row>
      {cols && (
        <div>
          <br />
          <Table striped bordered hover>
            <thead>
              <tr>
                <th>Display Name</th>
                <th>Description</th>
                <th>Required</th>
              </tr>
            </thead>
            <tbody>
              {cols.map((col: any) => (
                <tr>
                  <td>{col.displayName}</td>
                  <td>{col.description}</td>
                  <td>{col.required ? "True" : "False"}</td>
                </tr>
              ))}
            </tbody>
          </Table>
        </div>
      )}
    </div>
  ) : (
    <></>
  );
}
