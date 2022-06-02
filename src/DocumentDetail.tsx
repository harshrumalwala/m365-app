import { RouteComponentProps, useHistory, useParams } from "react-router-dom";
import { useAppContext } from "./AppContext";
import {
  deleteDoc,
  getDoc,
  getDocVersions,
  getListItem,
  searchSites,
  updateColumn,
  updateDoc,
} from "./GraphService";
import { useEffect, useState } from "react";
import { Button, Table, Form, Alert } from "react-bootstrap";
import "./DocumentDetail.css";

export default function DocumentDetail(props: RouteComponentProps) {
  const app = useAppContext();
  const history = useHistory();
  const { listItemId, siteName, listId, driveItemId } = useParams<{
    listItemId: string;
    siteName: string;
    listId: string;
    driveItemId: string;
  }>();
  const [fileDetail, setFileDetail] = useState<any>();
  const [name, setName] = useState<string>(fileDetail?.name);
  const [show, setShow] = useState<boolean>(false);
  const [fileVersions, setFileVersions] = useState<any>();
  const [refetchVersions, setRefetchVersions] = useState<boolean>(true);
  const [siteId, setSiteId] = useState<string>("");
  const [fields, setFields] = useState<any>({});
  const downloadUrl = fileDetail?.["@microsoft.graph.downloadUrl"];

  useEffect(() => {
    const getSiteDetails = async () => {
      const siteDetails = await searchSites(app.authProvider!, siteName);
      setSiteId(siteDetails[0].id);
    };
    getSiteDetails();
  }, [app.authProvider, siteName]);

  const getDocDetail = async () => {
    const doc = await getListItem(
      app.authProvider!,
      siteId,
      listId,
      listItemId
    );
    setFileDetail(doc);
  };

  useEffect(() => {
    const getDocVers = async () => {
      const docVers = await getDocVersions(
        app.authProvider!,
        siteId,
        listId,
        driveItemId
      );
      setFileVersions(docVers);
      setRefetchVersions(false);
    };
    if (listItemId && siteId) {
      getDocDetail();
      getDocVers();
    }
  }, [
    app.authProvider,
    listItemId,
    listId,
    refetchVersions,
    siteId,
    driveItemId,
  ]);

  useEffect(() => {
    setName(fileDetail?.name);
  }, [fileDetail]);

  const modifyDoc = async () => {
    const doc = await updateDoc(app.authProvider!, siteId, listItemId, {
      name: name,
    });
    if (doc?.id) {
      setShow(true);
      setRefetchVersions(true);
    }
  };

  const updateCol = async () => {
    const doc = await updateColumn(
      app.authProvider!,
      siteId,
      listId,
      listItemId,
      fields
    );
    if (doc) {
      getDocDetail();
      setShow(true);
      setRefetchVersions(true);
    }
  };

  const deleteFile = async () => {
    await deleteDoc(app.authProvider!, siteId, listItemId);
    history.push(`/sites/${siteName}/lists/${listId}/docs`);
  };
  console.log("fields", fields);
  return siteId !== "" ? (
    <div className="p-5 mb-4 bg-light rounded-3">
      <Alert
        show={show}
        variant="success"
        onClose={() => setShow(false)}
        dismissible
      >
        Your update was successful!
      </Alert>
      <div>
        {fileDetail && fileVersions && (
          <div>
            <div style={{ width: "100%", height: "60px" }}>
              <div className="left-button-container">
                <Button
                  variant="outline-dark"
                  onClick={() => {
                    window.open(fileDetail.webUrl);
                  }}
                >
                  Open in Sharepoint
                </Button>
                <Button
                  className="button-format"
                  variant="link"
                  onClick={() => {
                    window.open(downloadUrl);
                  }}
                >
                  Download
                </Button>
              </div>
              <div className="right-button-container">
                <Button className="button-format" onClick={updateCol}>
                  Update
                </Button>
                <Button
                  className="button-format"
                  variant="danger"
                  onClick={deleteFile}
                >
                  Delete
                </Button>
              </div>
            </div>
            {Object.entries(fileDetail.fields).map(
              (k: any) =>
                !k[0].includes("_") &&
                !k[0].includes("@") && (
                  <Form.Group className="mb-3">
                    <Form.Label>{k[0]}</Form.Label>
                    <Form.Control
                      value={k[0] in fields ? fields[k[0]] : k[1]}
                      onChange={(e: any) =>
                        setFields({ ...fields, [k[0]]: e.target.value })
                      }
                    />
                  </Form.Group>
                )
            )}
            <Form.Group className="mb-3">
              <Form.Label>Created By</Form.Label>
              <Form.Control
                value={fileDetail.createdBy.user.displayName}
                disabled
              />
            </Form.Group>
            <Form.Group className="mb-3">
              <Form.Label>Created Date</Form.Label>
              <Form.Control
                value={new Date(fileDetail.createdDateTime).toString()}
                disabled
              />
            </Form.Group>
            <Form.Group className="mb-3">
              <Form.Label>Last Modified By</Form.Label>
              <Form.Control
                value={fileDetail.lastModifiedBy.user.displayName}
                disabled
              />
            </Form.Group>
            <Form.Group className="mb-3">
              <Form.Label>Last Modified Date</Form.Label>
              <Form.Control
                value={new Date(fileDetail.lastModifiedDateTime).toString()}
                disabled
              />
            </Form.Group>
            <br />
            <h3>File Versions</h3>
            <br />
            <Table striped bordered hover>
              <thead>
                <tr>
                  <th>Version</th>
                  <th>Modified By</th>
                  <th>Modified Date </th>
                  <th>Size (B)</th>
                </tr>
              </thead>
              {fileVersions.map((v: any) => (
                <tbody>
                  <tr>
                    <td>{v.id}</td>
                    <td>{v.lastModifiedBy.user.displayName}</td>
                    <td>{new Date(v.lastModifiedDateTime).toString()}</td>
                    <td>{v.size}</td>
                  </tr>
                </tbody>
              ))}
            </Table>
          </div>
        )}
      </div>
    </div>
  ) : (
    <></>
  );
}
