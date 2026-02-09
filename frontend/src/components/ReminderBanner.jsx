import React, { useState, useEffect } from 'react';

const ReminderBanner = () => {
  const [message, setMessage] = useState(null);

  useEffect(() => {
    const checkReminder = () => {
      const now = new Date();
      const dayOfWeek = now.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday
      const date = now.getDate(); // 1-31

      // Only show on Mondays (1)
      if (dayOfWeek !== 1) {
        setMessage(null);
        return;
      }

      // Check if it is the first Monday of the month
      // If today is Monday and the date is <= 7, it must be the first Monday.
      const isFirstMonday = date <= 7;

      if (isFirstMonday) {
        setMessage("Reminder: Take Nizamabad-1 & Nizamabad-2 SEM data readings (including Time Drift)");
      } else {
        setMessage("Reminder: Take Nizamabad-1 & Nizamabad-2 SEM data readings");
      }
    };

    checkReminder();
    // Set up an interval to check periodically (e.g., every minute) in case the day changes while the app is open
    const interval = setInterval(checkReminder, 60000);

    return () => clearInterval(interval);
  }, []);

  if (!message) return null;

  return (
    <div className="flex-1 overflow-hidden mx-4 relative h-8 flex items-center">
      <style>
        {`
          @keyframes reminder-marquee {
            0% { transform: translateX(100%); }
            100% { transform: translateX(-100%); }
          }
          .animate-reminder-marquee {
            animation: reminder-marquee 20s linear infinite;
          }
          .animate-reminder-marquee:hover {
            animation-play-state: paused;
          }
          @media (prefers-reduced-motion: reduce) {
            .animate-reminder-marquee {
              animation: none;
              transform: none; /* Show static text if reduced motion is preferred */
              white-space: normal;
              text-align: center;
            }
          }
        `}
      </style>
      <div 
        className="w-full absolute whitespace-nowrap animate-reminder-marquee"
        role="status"
        aria-live="polite"
      >
        <span className="text-red-600 dark:text-red-400 font-bold text-lg animate-pulse drop-shadow-sm">
          {message}
        </span>
      </div>
    </div>
  );
};

export default ReminderBanner;
