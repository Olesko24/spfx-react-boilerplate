import * as React from 'react';
import '../../../styles/dist/tailwind.css';

export interface IAppProps {
  description: string;
}

const App = ({ description }: IAppProps): React.ReactElement => {
  return (
    <div className="w-full flex flex-col gap-4 text-center">
      <h1 className="text-3xl font-bold">Hello, {description}!</h1>
      <span className="text-lg">Willkommen beim SharePoint Template!</span>
      <span>Mit TailwindCSS v4!</span>
      <br />
      <span className='text-lg mb-2'>SharePoint / Fluent UI Theme Color Examples</span>
      <div className="text-neutral-foreground border">
        Should be your neutral foreground color
      </div>
      <button className="bg-brand-background hover:bg-brand-background-hover text-neutral-foreground-on-brand px-3 py-2 rounded">
        Brand button
      </button>
      <div className="bg-custom-bg text-custom-black p-3 rounded">
        custom bg + text
      </div>
    </div>
  );
};

export default App;
