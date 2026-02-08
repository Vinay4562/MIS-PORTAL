import { clsx } from "clsx";
import { twMerge } from "tailwind-merge"

export function cn(...inputs) {
  return twMerge(clsx(inputs));
}

export function formatDate(dateStr) {
  if (!dateStr) return "-";
  const parts = dateStr.split("-");
  if (parts.length === 3) {
    const [year, month, day] = parts;
    return `${day}-${month}-${year}`;
  }
  return dateStr;
}

export const downloadFile = async (data, filename) => {
  // Check if running in pywebview
  if (window.pywebview && window.pywebview.api) {
    try {
      // Convert Blob to Base64
      const reader = new FileReader();
      reader.readAsDataURL(new Blob([data]));
      reader.onloadend = async () => {
        const base64data = reader.result;
        const result = await window.pywebview.api.save_file(filename, base64data);
        if (result && result.success) {
            // Optional: notify success via return or callback, currently handling in UI
        }
      };
    } catch (error) {
      console.error("Pywebview download failed", error);
      throw error;
    }
  } else {
    // Standard Browser Download
    const url = window.URL.createObjectURL(new Blob([data]));
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', filename);
    document.body.appendChild(link);
    link.click();
    link.remove();
  }
};
