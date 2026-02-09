import React, { useState, useEffect } from 'react';
import axios from 'axios';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

const UnifiedReminderBanner = () => {
  const [mondayMessage, setMondayMessage] = useState(null);
  const [dailyMessages, setDailyMessages] = useState([]);

  // 1. Monday Reminder Logic
  useEffect(() => {
    const checkMondayReminder = () => {
      const now = new Date();
      const dayOfWeek = now.getDay(); // 0 = Sunday, 1 = Monday
      const date = now.getDate();

      // Only show on Mondays (1)
      if (dayOfWeek !== 1) {
        setMondayMessage(null);
        return;
      }

      const isFirstMonday = date <= 7;
      if (isFirstMonday) {
        setMondayMessage("Reminder: Take Nizamabad-1 & Nizamabad-2 SEM data readings (including Time Drift)");
      } else {
        setMondayMessage("Reminder: Take Nizamabad-1 & Nizamabad-2 SEM data readings");
      }
    };

    checkMondayReminder();
    const interval = setInterval(checkMondayReminder, 60000); // Check every minute
    return () => clearInterval(interval);
  }, []);

  // 2. Daily Status Logic
  useEffect(() => {
    const fetchDailyStatus = async () => {
      try {
        const token = localStorage.getItem('token');
        if (!token) return;

        const response = await axios.get(`${API}/daily-status`, {
          headers: { Authorization: `Bearer ${token}` }
        });
        
        const { line_losses, energy_consumption, max_min } = response.data;
        const newMessages = [];

        // Line Losses
        if (!line_losses.complete) {
            if (line_losses.missing_dates.length > 0) {
                const dates = line_losses.missing_dates.join(', ');
                newMessages.push(`Line Losses details are not captured for the last 2 days: ${dates}`);
            } else if (line_losses.missing_feeders.length > 0) {
                const feeders = line_losses.missing_feeders.join(', ');
                newMessages.push(`Pending Line Losses entries for today: ${feeders}`);
            }
        }

        // Energy Consumption
        if (!energy_consumption.complete) {
            if (energy_consumption.missing_dates.length > 0) {
                const dates = energy_consumption.missing_dates.join(', ');
                newMessages.push(`Energy Consumption details are not captured for the last 2 days: ${dates}`);
            } else if (energy_consumption.missing_sheets.length > 0) {
                 const sheets = energy_consumption.missing_sheets.join(', ');
                 newMessages.push(`Pending Energy Consumption entries for today: ${sheets}`);
            }
        }

        // Max Min
        if (!max_min.complete) {
            if (max_min.missing_dates.length > 0) {
                 const dates = max_min.missing_dates.join(', ');
                 newMessages.push(`Max–Min Data details are not captured for the last 2 days: ${dates}`);
            } else if (max_min.missing_feeders.length > 0) {
                 const feeders = max_min.missing_feeders.join(', ');
                 newMessages.push(`Pending Max–Min Data entries for today: ${feeders}`);
            }
        }
        
        setDailyMessages(newMessages);

      } catch (error) {
        console.error("Failed to fetch daily status:", error);
      }
    };

    fetchDailyStatus();
    const interval = setInterval(fetchDailyStatus, 5 * 60 * 1000); // Check every 5 mins
    return () => clearInterval(interval);
  }, []);

  // Combine messages
  const allMessages = [];
  if (mondayMessage) allMessages.push(mondayMessage);
  if (dailyMessages.length > 0) allMessages.push(...dailyMessages);

  if (allMessages.length === 0) return null;

  const combinedMessage = allMessages.join(" | ");

  return (
    <div className="flex-1 overflow-hidden mx-4 relative h-8 flex items-center">
      <style>
        {`
          @keyframes unified-marquee {
            0% { transform: translateX(100vw); }
            100% { transform: translateX(-100%); }
          }
          @keyframes glow-pulse {
            0%, 100% { filter: drop-shadow(0 0 2px rgba(220, 38, 38, 0.5)); opacity: 1; }
            50% { filter: drop-shadow(0 0 8px rgba(220, 38, 38, 1)); opacity: 0.8; }
          }
          .animate-unified-marquee {
            animation: unified-marquee ${Math.max(20, combinedMessage.length * 0.15)}s linear infinite;
          }
          .animate-glow-pulse {
            animation: glow-pulse 2s ease-in-out infinite;
          }
          .banner-container:hover .animate-unified-marquee,
          .banner-container:hover .animate-glow-pulse {
            animation-play-state: paused;
          }
        `}
      </style>
      <div 
        className="absolute whitespace-nowrap animate-unified-marquee banner-container"
        role="status"
        aria-live="polite"
      >
        <span className="text-red-600 dark:text-red-400 font-bold text-lg animate-glow-pulse inline-block">
          {combinedMessage}
        </span>
      </div>
    </div>
  );
};

export default UnifiedReminderBanner;
