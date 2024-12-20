import React from 'react';
import { Button, Box } from '@mui/material';
import { ChevronLeft, ChevronRight } from '@mui/icons-material';

const NavigationControls = ({ onPrevious, onNext }) => {
  return (
    <Box sx={{ display: 'flex', justifyContent: 'center', gap: 2, mb: 2 }}>
      <Button 
        startIcon={<ChevronLeft />} 
        onClick={onPrevious}
        variant="contained"
      >
        Previous Week
      </Button>
      <Button 
        endIcon={<ChevronRight />} 
        onClick={onNext}
        variant="contained"
      >
        Next Week
      </Button>
    </Box>
  );
};

export default NavigationControls; 