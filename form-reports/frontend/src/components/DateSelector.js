import React from 'react';
import { TextField } from '@mui/material';
import { format } from 'date-fns';

const DateSelector = ({ currentDate, onDateChange }) => {
  return (
    <TextField
      type="date"
      value={format(currentDate, 'yyyy-MM-dd')}
      onChange={(e) => onDateChange(new Date(e.target.value))}
      sx={{ mb: 2 }}
    />
  );
};

export default DateSelector; 