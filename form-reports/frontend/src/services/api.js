import axios from 'axios';

const API_BASE_URL = process.env.REACT_APP_API_URL || 'http://localhost:5000';

export const fetchTimeEntries = async (startDate, endDate) => {
  try {
    const response = await axios.get(`${API_BASE_URL}/time-entries`, {
      params: {
        startDate: startDate.toISOString(),
        endDate: endDate.toISOString()
      }
    });
    return response.data;
  } catch (error) {
    console.error('Error fetching time entries:', error);
    return [];
  }
}; 