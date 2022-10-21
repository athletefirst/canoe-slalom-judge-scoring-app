import * as React from "react";
import Stack from "@mui/material/Stack";
import Typography from "@mui/material/Typography";
import Button from '@mui/material/Button';
import { grey, green, red, blue, } from '@mui/material/colors';
import { useData } from "../providers/CanoeSlalomHeatDataProvider";

function GateNumber({ num, isDownStream = true, isLocked = false }) {
    const gateSx = {
        fontSize: '1.3rem',
        width: '42px',
        height: '42px',
        lineHeight: '42px',
        border: 1,
        borderRadius: '50%',
        borderColor: isLocked ? grey[200] : isDownStream ? green[500] : red[500],
        color: isLocked ? grey[400] : blue[700],
        userSelect: 'none',
    };
    return (
        <Typography variant="caption" sx={gateSx} align="center">
            {num}
        </Typography>
    );
}

function JudgeButton({ judge, isError = false, isFailure = false, isLoading = false, isLocked = false, onSelectedJudge = f => f }) {
    const onSelectedJudge0 = () => onSelectedJudge('0');
    const onSelectedJudge2 = () => onSelectedJudge('2');
    const onSelectedJudge50 = () => onSelectedJudge('50');
    const onSelectedJudgeDsq = () => onSelectedJudge('DSQ');

    const color = isLoading ? 'secondary' : isFailure ? 'warning' : isError ? 'error' : 'primary'
    const sx = {
        width: "64px",
        height: "44px",
        fontSize: "1.2rem",
    };
    const sxLeft = {
        ...sx,
        borderTopRightRadius: 0,
        borderBottomRightRadius: 0,
    };
    const sxMiddle = {
        ...sx,
        borderRadius: 0,
    };
    const sxRight = {
        ...sx,
        borderTopLeftRadius: 0,
        borderBottomLeftRadius: 0,
    };
    return (
        <Stack direction="row" spacing={'-1px'} sx={{ m: '0px' }}>
            <Button variant={judge === '0' ? 'contained' : 'outlined'} disabled={isLocked} color={color} sx={sxLeft} onClick={onSelectedJudge0}>0</Button>
            <Button variant={judge === '2' ? 'contained' : 'outlined'} disabled={isLocked} color={color} sx={sxMiddle} onClick={onSelectedJudge2}>2</Button>
            <Button variant={judge === '50' ? 'contained' : 'outlined'} disabled={isLocked} color={color} sx={sxMiddle} onClick={onSelectedJudge50}>50</Button>
            <Button variant={judge === 'DSQ' ? 'contained' : 'outlined'} disabled={isLocked} color={color} sx={sxRight} onClick={onSelectedJudgeDsq}>DSQ</Button>
        </Stack>
    );
}

export default function MuiGateJudgeButton({ row, bib, tag, num, judge, isDownStream = true, isError = false, isFailure = false, isLoading = false, isLocked = false }) {
    const { setGateJudge } = useData();
    const onSelectedJudge = judge => {
        if ((!isLoading) && (!isLocked)) {
            setGateJudge(row, num, judge);
        }
    };

    return (
        <Stack direction="row" spacing={1} sx={{ m: '0px' }}>
            <GateNumber num={num} isDownStream={isDownStream} isLocked={isLocked} />
            <JudgeButton judge={judge} isError={isError} isFailure={isFailure} isLoading={isLoading} isLocked={isLocked} onSelectedJudge={onSelectedJudge} />
        </Stack>
    );
}