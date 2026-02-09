import React, { useState, useEffect } from 'react';
import axios from 'axios';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

const DailyReminderBanner = () => {
  const [messages, setMessages] = useState([]);

  useEffect(() => {
    const fetchStatus = async () => {
      try {
        const token = localStorage.getItem('token');
        if (!token) return;

        const response = await axios.get(`${API}/daily-status`, {
          headers: { Authorization: `Bearer ${token}` }
        });
        
        const { line_losses, energy_consumption, max_min } = response.data;
        const newMessages = [];

        // 1. Line Losses
        if (!line_losses.complete) {
            if (line_losses.missing_dates.length > 0) {
                const dates = line_losses.missing_dates.join(', ');
                newMessages.push(`Line Losses details are not captured for the last 2 days: ${dates}`);
            } else if (line_losses.missing_feeders.length > 0) {
                const feeders = line_losses.missing_feeders.join(', ');
                newMessages.push(`Pending Line Losses entries for today: ${feeders}`);
            }
        }

        // 2. Energy Consumption
        if (!energy_consumption.complete) {
            if (energy_consumption.missing_dates.length > 0) {
                const dates = energy_consumption.missing_dates.join(', ');
                newMessages.push(`Energy Consumption details are not captured for the last 2 days: ${dates}`);
            } else if (energy_consumption.missing_sheets.length > 0) {
                 const sheets = energy_consumption.missing_sheets.join(', ');
                 newMessages.push(`Pending Energy Consumption entries for today: ${sheets}`);
            }
        }

        // 3. Max Min
        if (!max_min.complete) {
            if (max_min.missing_dates.length > 0) {
                 const dates = max_min.missing_dates.join(', ');
                 newMessages.push(`Max–Min Data details are not captured for the last 2 days: ${dates}`);
            } else if (max_min.missing_feeders.length > 0) {
                 const feeders = max_min.missing_feeders.join(', ');
                 newMessages.push(`Pending Max–Min Data entries for today: ${feeders}`);
            }
        }
        
        setMessages(newMessages);

      } catch (error) {
        console.error("Failed to fetch daily status:", error);
      }
    };

    fetchStatus();
    // Check every 5 minutes
    const interval = setInterval(fetchStatus, 5 * 60 * 1000); 
    return () => clearInterval(interval);
  }, []);

  if (messages.length === 0) return null;

  const combinedMessage = messages.join(" | ");

  return (
    <div className="flex-1 overflow-hidden mx-4 relative h-8 flex items-center">
      <style>
        {`
          @keyframes daily-reminder-marquee {
            0% { transform: translateX(100%); }
            100% { transform: translateX(-100%); }
          }
          .animate-daily-reminder-marquee {
            animation: daily-reminder-marquee ${Math.max(20, combinedMessage.length * 0.2)}s linear infinite;
          }
          .animate-daily-reminder-marquee:hover {
            animation-play-state: paused;
          }
        `}
      </style>
      <div 
        className="w-full absolute whitespace-nowrap animate-daily-reminder-marquee"
        role="status"
        aria-live="polite"
      >
        <span className="text-red-600 dark:text-red-400 font-bold text-lg drop-shadow-sm">
          {combinedMessage}
        </span>
      </div>
    </div>
  );
};

export default DailyReminderBanner;
