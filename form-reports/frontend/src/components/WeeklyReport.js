import React, { useState, useEffect } from 'react';
import { format, addWeeks, startOfWeek, endOfWeek } from 'date-fns';
import { Paper, Container, Typography } from '@mui/material';
import TimeEntry from './TimeEntry';
import DateSelector from './DateSelector';
import NavigationControls from './NavigationControls';
import { fetchTimeEntries } from '../services/api';

const WeeklyReport = () => {
  const [entries, setEntries] = useState([]);
  const [currentDate, setCurrentDate] = useState(new Date());

  useEffect(() => {
    const startDate = startOfWeek(currentDate, { weekStartsOn: 1 }); // Monday
    const endDate = endOfWeek(currentDate, { weekStartsOn: 1 }); // Sunday
    
    const loadEntries = async () => {
      const data = await fetchTimeEntries(startDate, endDate);
      setEntries(data);
    };

    loadEntries();
  }, [currentDate]);

  const handlePreviousWeek = () => {
    setCurrentDate(date => addWeeks(date, -1));
  };

  const handleNextWeek = () => {
    setCurrentDate(date => addWeeks(date, 1));
  };

  const handleDateChange = (date) => {
    setCurrentDate(date);
  };

  return (
    <Container maxWidth="lg">
      <Typography variant="h4" align="center" gutterBottom>
        Weekly Activity Report
      </Typography>
      
      <DateSelector 
        currentDate={currentDate} 
        onDateChange={handleDateChange}
      />
      
      <NavigationControls
        onPrevious={handlePreviousWeek}
        onNext={handleNextWeek}
      />

      {entries.map(entry => (
        <TimeEntry key={entry.id} entry={entry} />
      ))}
    </Container>
  );
};

export default WeeklyReport; 