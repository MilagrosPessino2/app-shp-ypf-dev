import * as React from "react";
import { Routes, Route } from "react-router-dom";
import Forbidden from "../forbidden/Forbidden";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import NovedadesPage from "../pages/NovedadesPage";

interface IAppProps {
    context: WebPartContext;
}

const App: React.FC<IAppProps> = ({ context }) => {
    return (
        <Routes>
            <Route path="/" element={<NovedadesPage context={context} />} />
            <Route path="/forbidden" element={<Forbidden />} />
        </Routes>
    );
};

export default App;
