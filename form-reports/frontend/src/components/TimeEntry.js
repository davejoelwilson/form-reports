import React from 'react';
import { Paper, Grid, Typography } from '@mui/material';

const TimeEntry = ({ entry }) => {
  return (
    <Paper sx={{ p: 2, mb: 2, backgroundColor: '#fff' }}>
      <Grid container spacing={2}>
        <Grid item xs={2}>
          <Typography variant="subtitle2" sx={{ backgroundColor: '#ccffcc', p: 1 }}>Date</Typography>
        </Grid>
        <Grid item xs={10}>
          <Typography>{entry.date}</Typography>
        </Grid>

        <Grid item xs={2}>
          <Typography variant="subtitle2" sx={{ backgroundColor: '#ccffcc', p: 1 }}>Start Time</Typography>
        </Grid>
        <Grid item xs={10}>
          <Typography>{entry.startTime}</Typography>
        </Grid>

        {/* Add other fields similar to the Word template */}
      </Grid>
    </Paper>
  );
};

export default TimeEntry; 