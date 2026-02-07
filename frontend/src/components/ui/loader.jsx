import { Loader2 } from 'lucide-react';
import { cn } from '@/lib/utils';

export function Loader({ className, size = "default", ...props }) {
  const sizeClasses = {
    sm: "h-4 w-4",
    default: "h-8 w-8",
    lg: "h-12 w-12",
    xl: "h-16 w-16"
  };

  return (
    <Loader2 
      className={cn("animate-spin text-blue-600 dark:text-blue-400", sizeClasses[size], className)} 
      {...props} 
    />
  );
}

export function FullPageLoader({ text = "Loading..." }) {
  return (
    <div className="fixed inset-0 bg-white/80 dark:bg-slate-950/80 backdrop-blur-sm z-50 flex flex-col items-center justify-center">
      <Loader size="xl" className="mb-4" />
      <p className="text-lg font-medium text-slate-700 dark:text-slate-300 animate-pulse">
        {text}
      </p>
    </div>
  );
}

export function BlockLoader({ text = "Loading...", className }) {
  return (
    <div className={cn("absolute inset-0 bg-white/60 dark:bg-slate-950/60 backdrop-blur-[1px] z-10 flex flex-col items-center justify-center min-h-[200px]", className)}>
      <Loader size="lg" className="mb-3" />
      <p className="text-sm font-medium text-slate-600 dark:text-slate-400">
        {text}
      </p>
    </div>
  );
}

export function InlineLoader({ text, className }) {
  return (
    <div className={cn("flex items-center justify-center p-4", className)}>
      <Loader size="default" className="mr-2" />
      {text && <span className="text-sm text-slate-600 dark:text-slate-400">{text}</span>}
    </div>
  );
}
