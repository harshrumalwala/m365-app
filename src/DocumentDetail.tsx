import { RouteComponentProps, useParams } from "react-router-dom";
import { useAppContext } from "./AppContext";
import { AsyncTypeahead, Menu } from "react-bootstrap-typeahead";
import { getDoc, searchDocs, uploadDocument } from "./GraphService";
import { useEffect, useState } from "react";
import { Button } from "react-bootstrap";
import "./DocumentDetail.css";

export default function DocumentDetail(props: RouteComponentProps) {
  const app = useAppContext();
  const { id } = useParams<{ id: string }>();
  const [fileDetail, setFileDetail] = useState<any>();
  const downloadUrl = fileDetail?.["@microsoft.graph.downloadUrl"];

  useEffect(() => {
    const getDocDetail = async () => {
      const doc = await getDoc(app.authProvider!, id);
      setFileDetail(doc);
    };
    if (id) {
      getDocDetail();
    }
  }, [app.authProvider, id]);

  return (
    <div className="p-5 mb-4 bg-light rounded-3">
      <div>
        <h2>File Details</h2>
        {fileDetail && (
          <div>
            <p>File Name : {fileDetail.name}</p>
            <p>Created By : {fileDetail.createdBy.user.displayName}</p>
            <p>
              Created Date : {new Date(fileDetail.createdDateTime).toString()}
            </p>
            <p>
              Last Modified By : {fileDetail.lastModifiedBy.user.displayName}
            </p>
            <p>
              Last Modified Date :{" "}
              {new Date(fileDetail.lastModifiedDateTime).toString()}
            </p>
            <div>
              <Button
                className="button-format"
                onClick={() => {
                  window.open(downloadUrl);
                }}
              >
                Download
              </Button>
              <Button
                className="button-format"
                variant="outline-dark"
                onClick={() => {
                  window.open(fileDetail.webUrl);
                }}
              >
                Open in Sharepoint
              </Button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
