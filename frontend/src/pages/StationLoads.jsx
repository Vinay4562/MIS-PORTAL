import { Button } from '@/components/ui/button';

const StationLoads = () => {
  return (
    <div className="flex flex-col h-full">
      <header className="bg-white/90 dark:bg-slate-900/90 backdrop-blur-xl border-b border-slate-200/70 dark:border-slate-800/70 px-4 md:px-8 py-4 flex items-center justify-between">
        <h2 className="text-lg font-semibold text-slate-900 dark:text-slate-50">
          All the feeders data of 400KV, 220KV and ICT’s and Station Loads in MVA and MW.
        </h2>
      </header>
      <div className="flex-1 flex items-center justify-center">
        <a href="https://neon-feeder-flow.vercel.app/" target="_blank" rel="noopener noreferrer">
          <Button>Access Data</Button>
        </a>
      </div>
    </div>
  );
};

export default StationLoads;
