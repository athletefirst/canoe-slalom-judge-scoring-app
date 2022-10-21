import * as React from "react";
import Stack from "@mui/material/Stack";
import Typography from "@mui/material/Typography";
import Button from '@mui/material/Button';
import TextField from '@mui/material/TextField';
import { styled, useTheme } from '@mui/material/styles';
import { grey, green, red, blue, } from '@mui/material/colors';
import { useData } from "../providers/CanoeSlalomHeatDataProvider";
import MuiTimeInputDialog, { TimeJudgeData } from './MuiTimeInputDialog';
import CanoeSlalomHeatData from '../../dao/CanoeSlalomHeatData';

function TimeLabel({ startOrFinish, isLocked = false }) {
    const sx = {
        fontSize: '1.3rem',
        width: '42px',
        height: '42px',
        lineHeight: '42px',
        border: 1,
        borderRadius: '3px',
        borderColor: isLocked ? grey[200] : blue[500],
        color: isLocked ? grey[400] : blue[700],
        userSelect: 'none',
    };
    const caption = startOrFinish === 'START' ? 'S' : 'F'
    return (
        <Typography variant="caption" sx={sx} align="center">
            {caption}
        </Typography>
    );
}

function TimeField({ bib, tag, startOrFinish, seconds, judge, isError = false, isFailure = false, isLoading = false, isLocked = false, onChangeTimeJudge = (s, j) => { } }) {

    const hms = CanoeSlalomHeatData.secondsToHms(seconds);
    const hhPart = hms.hours > 0 ? String(hms.hours) + ':' : '';
    const mmPart = hhPart.length > 0 ? ('0' + String(hms.minutes)).slice(-2) + ':' : hms.minutes > 0 ? String(hms.minutes) + ':' : '';
    const ssPart = mmPart.length > 0 ? ('0' + hms.seconds.toFixed(3)).slice(-6) : hms.seconds > 0 ? hms.seconds.toFixed(3) : '-.---';
    const timeText = hhPart + mmPart + ssPart;

    let hh: number | undefined = undefined;
    let mm: number | undefined = undefined;
    let ss: number | undefined = undefined;
    if (hms.hours > 0) {
        hh = hms.hours;
    }
    if ((hms.hours > 0) || (hms.minutes > 0)) {
        mm = hms.minutes;
    }
    if ((hms.hours > 0) || (hms.minutes > 0) || (hms.seconds > 0)) {
        ss = hms.seconds;
    }
    const originTimeJudgeData: TimeJudgeData = {
        hh,
        mm,
        ss,
        judge,
    };

    const [timeJudgeData, setTimeJudgeData] = React.useState(originTimeJudgeData);
    const [open, setOpen] = React.useState(false);

    const handleEdit = () => {
        if (isLocked) {
            return;
        }
        setTimeJudgeData(originTimeJudgeData);
        setOpen(true);
    };

    const onClose = (apply: boolean) => {
        setOpen(false);
        if (apply) {
            const seconds = (timeJudgeData.hh ? Number(timeJudgeData.hh) * 3600 : 0)
                + (timeJudgeData.mm ? Number(timeJudgeData.mm) * 60 : 0)
                + (timeJudgeData.ss ? Number(timeJudgeData.ss) : 0);
            onChangeTimeJudge(seconds, timeJudgeData.judge);
        }
    };

    const color = isLoading ? 'secondary' : isFailure ? 'warning' : isError ? 'error' : 'primary'

    let buttonVariant: ("contained" | "outlined") = 'contained';
    let judgeText: ('ENT' | '+50' | 'DNS' | 'DNF' | 'DSQ' | '---');
    switch (judge) {
        case 'STARTED':
            judgeText = 'ENT';
            break;
        case 'FINISHED':
            judgeText = 'ENT';
            break;
        case 'FINISHED_50':
            judgeText = '+50';
            break;
        case 'DNS':
            judgeText = 'DNS';
            buttonVariant = 'outlined';
            break;
        case 'DNF':
            judgeText = 'DNF';
            buttonVariant = 'outlined';
            break;
        case 'DSQ':
            judgeText = 'DSQ';
            buttonVariant = 'outlined';
            break;
        default:
            judgeText = '---';
            buttonVariant = 'outlined';
            break;
    }

    const ConvColor = (color: string) => {
        const theme = useTheme();
        const palette: any = theme.palette;
        return palette[color].main;
    };
    const colorCode = ConvColor(color);

    const CssTextField = styled(TextField)({
        '& label.Mui-focused': {
            color: colorCode,
        },
        '& .MuiOutlinedInput-root': {
            '&.Mui-focused fieldset': {
                borderColor: colorCode,
                borderWidth: '1px',
            },
        },
    });

    const textSx = {
        width: '186px',
        paddingTop: '1px',
        input: {
            height: '40px',
            textAlign: 'right',
            fontSize: '1.6rem',
            paddingLeft: '0px',
            paddingRight: '10px',
            paddingTop: '3px',
            paddingBottom: '0px',
            color: colorCode,
        },
    };
    const textInputProps = {
        readOnly: true,
    };

    const buttonSx = {
        width: '64px',
        height: '44px',
        fontSize: '1.2rem',
    };

    return (
        <>
            <Stack direction="row" spacing={"3px"} sx={{ m: '0px' }}>
                <CssTextField sx={textSx} disabled={isLocked} focused InputProps={textInputProps} value={timeText} label="hh:mm:ss.SSS" onClick={handleEdit} />
                <Button sx={buttonSx} disabled={isLocked} color={color} variant={buttonVariant} onClick={handleEdit} >{judgeText}</Button>
            </Stack>
            <MuiTimeInputDialog open={open} bib={bib} tag={tag} startOrFinish={startOrFinish} timeJudgeData={timeJudgeData} onChange={setTimeJudgeData} onClose={onClose} />
        </>
    );
}

export default function MuiTimeInput({ row, bib, tag, startOrFinish, seconds, judge, isError = false, isFailure = false, isLoading = false, isLocked = false }) {
    const { setStartedTime, setFinishedTime } = useData();
    const onChangeTimeJudge = (seconds, judge) => {
        if ((!isLoading) && (!isLocked)) {
            if (startOrFinish === 'START') {
                setStartedTime(row, seconds, judge);
            } else {
                setFinishedTime(row, seconds, judge);
            }
        }
    };

    return (
        <Stack direction="row" spacing={1} sx={{ m: '0px' }}>
            <TimeLabel startOrFinish={startOrFinish} isLocked={isLocked} />
            <TimeField bib={bib} tag={tag} startOrFinish={startOrFinish} seconds={seconds} judge={judge} isError={isError} isFailure={isFailure} isLoading={isLoading} isLocked={isLocked} onChangeTimeJudge={onChangeTimeJudge} />
        </Stack>
    );
}