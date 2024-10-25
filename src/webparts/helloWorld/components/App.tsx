import * as React from 'react';

export interface IAppProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}

const App = ({ description }: IAppProps) => {
  return (
    <div className="w-full flex flex-col gap-4">
      <h1 className="text-3xl font-bold text-center">Hello, {description}!</h1>
      <h2>Willkommen beim SharePoint Template!</h2>
    </div>
  );
};

export default App;
