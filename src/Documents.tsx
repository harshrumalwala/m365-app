import { Link, RouteComponentProps, useParams } from "react-router-dom";
import { useAppContext } from "./AppContext";
import { AsyncTypeahead, Menu } from "react-bootstrap-typeahead";
import { searchDocs, searchSites, uploadDocument } from "./GraphService";
import { useEffect, useState } from "react";
import "./Documents.css";
import { Alert, Button, Form } from "react-bootstrap";

export default function Documents(props: RouteComponentProps) {
  const { siteName } = useParams<{ siteName: string }>();
  const app = useAppContext();
  const [matchedDocs, setMatchedDocs] = useState<Array<any>>([]);
  const [show, setShow] = useState<boolean>(false);
  const [siteId, setSiteId] = useState<string>("");
  const [fileDetails, setFileDetails] = useState<{
    name: string;
    type: string;
    lastModifiedDate: Date;
  }>();

  useEffect(() => {
    const getSiteDetails = async () => {
      const siteDetails = await searchSites(app.authProvider!, siteName);
      setSiteId(siteDetails[0].id);
    };
    getSiteDetails();
  }, [app.authProvider, siteName]);

  const fileUpload = () => {
    const filename = fileDetails?.name;
    const filereader = new FileReader();
    filereader.onload = async (event) => {
      const resp = await uploadDocument(
        app.authProvider!,
        siteId,
        filename as string,
        fileDetails
      );
      if (resp?.id) setShow(true);
    };
    filereader.readAsArrayBuffer(fileDetails as unknown as Blob);
  };

  return siteId !== "" ? (
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
          const docs = await searchDocs(app.authProvider!, siteId, query);
          setMatchedDocs(docs);
        }}
        options={matchedDocs}
        renderMenu={(results, menuProps) => (
          <Menu {...menuProps}>
            {results.map((option: any) => (
              <div className="select-item">
                <Link to={`/sites/${siteName}/docs/${option.id}`}>
                  {option.name}
                </Link>
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
  ) : (
    <></>
  );
}
