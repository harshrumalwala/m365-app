// contexts/User/index.jsx

import React, {
  Dispatch,
  FC,
  SetStateAction,
  useContext,
  useState,
} from "react";

export const SiteContext = React.createContext({
  siteId: "",
  setSiteId: (() => null) as unknown as Dispatch<SetStateAction<string>>,
});

export const SiteProvider: FC = ({ children }) => {
  const [siteId, setSiteId] = useState<string>("");
  return (
    <SiteContext.Provider value={{ siteId, setSiteId }}>
      {children}
    </SiteContext.Provider>
  );
};

export const useCurrentSite = () => useContext(SiteContext);
