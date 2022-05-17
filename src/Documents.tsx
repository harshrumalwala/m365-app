import { RouteComponentProps } from "react-router-dom";
import { useAppContext } from "./AppContext";
import { AsyncTypeahead, Menu } from "react-bootstrap-typeahead";
import { searchDocs, uploadDocument } from "./GraphService";
import { useState } from "react";
import "./Documents.css";

export default function Documents(props: RouteComponentProps) {
  const app = useAppContext();
  const [matchedDocs, setMatchedDocs] = useState<Array<any>>([]);
  const [fileDetails, setFileDetails] = useState<{
    name: string;
    type: string;
    lastModifiedDate: Date;
  }>();

  const openDoc = (doc: any) => {
    window.open(doc.webUrl);
  };

  const fileUpload = () => {
    const filename = fileDetails?.name;
    const filereader = new FileReader();
    filereader.onload = async (event) => {
      await uploadDocument(filename as string, fileDetails);
    };
    filereader.readAsArrayBuffer(fileDetails as unknown as Blob);
  };

  return (
    <div className="p-5 mb-4 bg-light rounded-3">
      <AsyncTypeahead
        isLoading={false}
        filterBy={() => true}
        onSearch={async (query) => {
          const docs = await searchDocs(app.authProvider!, query);
          console.log("docs", docs);
          setMatchedDocs(docs);
        }}
        options={matchedDocs}
        renderMenu={(results, menuProps) => (
          <Menu {...menuProps}>
            {results.map((option: any) => (
              <div className="select-item" onClick={() => openDoc(option)}>
                {option.name}
              </div>
            ))}
          </Menu>
        )}
      />
      <div className="upload-container">
        <h4>Upload File To Sharepoint</h4>
        <div>
          <input
            type="file"
            onChange={(e: any) => setFileDetails(e.target.files[0])}
          />
          <button onClick={() => fileUpload()}>Upload!</button>
        </div>
      </div>
    </div>
  );
}
