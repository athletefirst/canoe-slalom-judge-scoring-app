import * as React from 'react';
import AppBar from '@mui/material/AppBar';
import Box from '@mui/material/Box';
import Toolbar from '@mui/material/Toolbar';
import Typography from '@mui/material/Typography';
import KayakingIcon from '@mui/icons-material/Kayaking';
import IconButton from '@mui/material/IconButton';
import MenuIcon from '@mui/icons-material/Menu';
import RefreshIcon from '@mui/icons-material/Refresh';
import CircularProgress from '@mui/material/CircularProgress';
import { useData } from "../providers/CanoeSlalomHeatDataProvider";

export default function DatasetAppBar() {
    const { error, loading, dataset, loadDataset } = useData();
    return (
        <Box className="AppBar">
            <AppBar position="static" color={error ? 'secondary' : 'primary'}>
                <Toolbar>
                    <KayakingIcon fontSize="large" />
                    <Typography variant="h6" component="div" sx={{ flexGrow: 1, }}>
                        {dataset.heatName}
                    </Typography>
                    {loading ?
                        <Box color="gray.500" width="51px" height="51px" display="flex" justifyContent="center" alignItems="center">
                            <CircularProgress color="inherit" size="27px" />
                        </Box>
                        :
                        <IconButton color="inherit" onClick={loadDataset}>
                            <RefreshIcon fontSize="large" />
                        </IconButton>}
                </Toolbar>
            </AppBar>
        </Box>
    );
}