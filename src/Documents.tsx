import { Link, RouteComponentProps } from "react-router-dom";
import { useAppContext } from "./AppContext";
import { AsyncTypeahead, Input, Menu } from "react-bootstrap-typeahead";
import { searchDocs, uploadDocument } from "./GraphService";
import { useState } from "react";
import "./Documents.css";
import { Alert, Button, Form } from "react-bootstrap";

export default function Documents(props: RouteComponentProps) {
  const app = useAppContext();
  const [matchedDocs, setMatchedDocs] = useState<Array<any>>([]);
  const [show, setShow] = useState<boolean>(false);
  const [fileDetails, setFileDetails] = useState<{
    name: string;
    type: string;
    lastModifiedDate: Date;
  }>();

  const fileUpload = () => {
    const filename = fileDetails?.name;
    const filereader = new FileReader();
    filereader.onload = async (event) => {
      const resp = await uploadDocument(
        app.authProvider!,
        filename as string,
        fileDetails
      );
      if (resp?.id) setShow(true);
    };
    filereader.readAsArrayBuffer(fileDetails as unknown as Blob);
  };

  return (
    <div className="p-5 mb-4 bg-light rounded-3">
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
        placeholder="Search in Sharepoint"
        filterBy={() => true}
        onSearch={async (query) => {
          const docs = await searchDocs(app.authProvider!, query);
          setMatchedDocs(docs);
        }}
        options={matchedDocs}
        renderMenu={(results, menuProps) => (
          <Menu {...menuProps}>
            {results.map((option: any) => (
              <div className="select-item">
                <Link to={`/docs/${option.id}`}>{option.name}</Link>
              </div>
            ))}
          </Menu>
        )}
      />
      <div className="upload-container">
        <div>
          <Form.Group controlId="formFile" className="mb-3">
            <Form.Label>Upload File To Sharepoint</Form.Label>
            <Form.Control
              type="file"
              onChange={(e: any) => setFileDetails(e.target.files[0])}
            />
          </Form.Group>
          <Button className="button-format" onClick={() => fileUpload()}>
            Upload
          </Button>
        </div>
      </div>
    </div>
  );
}
