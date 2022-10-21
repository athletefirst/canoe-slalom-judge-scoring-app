import * as React from "react";
import Dialog from "@mui/material/Dialog";
import DialogTitle from "@mui/material/DialogTitle";
import DialogContent from "@mui/material/DialogContent";
import DialogActions from "@mui/material/DialogActions";
import Stack from "@mui/material/Stack";
import Typography from "@mui/material/Typography";
import TextField from "@mui/material/TextField";
import Button from "@mui/material/Button";
import NumberFormat from "react-number-format";
import { formControlLabelClasses } from "@mui/material";

export type TimeJudgeData = {
    hh: any;
    mm: any;
    ss: any;
    judge: any;
};

type TimeJudgeDialogProps = {
    open: boolean;
    bib: string;
    tag: string;
    startOrFinish: string;
    timeJudgeData: TimeJudgeData;
    onChange: (values: TimeJudgeData) => void;
    onClose: (apply: boolean) => void;
};

function JudgeButton({
    startOrFinish,
    judge,
    isError = false,
    isFailure = false,
    isLoading = false,
    isLocked = false,
    onSelected = (f) => f
}) {
    const onSelectedStarted = () => onSelected("STARTED");
    const onSelectedDns = () => onSelected("DNS");
    const onSelectedFinished = () => onSelected("FINISHED");
    const onSelectedFinished50 = () => onSelected("FINISHED_50");
    const onSelectedDnf = () => onSelected("DNF");
    const onSelectedDsq = () => onSelected("DSQ");

    const color = isLoading
        ? "secondary"
        : isFailure
            ? "warning"
            : isError
                ? "error"
                : "primary";
    const sx = {
        width: "64px",
        height: "42px",
        fontSize: "1.2rem"
    };
    const sxLeft = {
        ...sx,
        borderTopRightRadius: 0,
        borderBottomRightRadius: 0
    };
    const sxMiddle = {
        ...sx,
        borderRadius: 0
    };
    const sxRight = {
        ...sx,
        borderTopLeftRadius: 0,
        borderBottomLeftRadius: 0
    };

    const StartedButton = () => (
        <Button
            variant={judge === "STARTED" ? "contained" : "outlined"}
            disabled={isLocked}
            color={color}
            sx={{ ...sxLeft, width: "127px" }}
            onClick={onSelectedStarted}
        >
            ENT
        </Button>
    );
    const FinishedButton = () => (
        <Button
            variant={judge === "FINISHED" ? "contained" : "outlined"}
            disabled={isLocked}
            color={color}
            sx={sxLeft}
            onClick={onSelectedFinished}
        >
            ENT
        </Button>
    );
    const Finished50Button = () => (
        <Button
            variant={judge === "FINISHED_50" ? "contained" : "outlined"}
            disabled={isLocked}
            color={color}
            sx={sxMiddle}
            onClick={onSelectedFinished50}
        >
            +50
        </Button>
    );
    const DnsButton = () => (
        <Button
            variant={judge === "DNS" ? "contained" : "outlined"}
            disabled={isLocked}
            color={color}
            sx={sxMiddle}
            onClick={onSelectedDns}
        >
            DNS
        </Button>
    );
    const DnfButton = () => (
        <Button
            variant={judge === "DNF" ? "contained" : "outlined"}
            disabled={isLocked}
            color={color}
            sx={sxMiddle}
            onClick={onSelectedDnf}
        >
            DNF
        </Button>
    );
    const DsqButton = () => (
        <Button
            variant={judge === "DSQ" ? "contained" : "outlined"}
            disabled={isLocked}
            color={color}
            sx={sxRight}
            onClick={onSelectedDsq}
        >
            DSQ
        </Button>
    );

    return (
        <Stack spacing={"-1px"} direction="row" sx={{ m: "0px" }}>
            {startOrFinish === "START" ? (
                <StartedButton />
            ) : (
                <>
                    <FinishedButton />
                    <Finished50Button />
                </>
            )}
            {startOrFinish === "START" ? (
                <>
                    <DnsButton />
                    <DsqButton />
                </>
            ) : (
                <>
                    <DnfButton />
                    <DsqButton />
                </>
            )}
        </Stack>
    );
}

function TimeCaption() {
    const sx = {
        fontSize: "1rem",
        width: "55px",
        height: "65px",
        lineHeight: "74px",
        userSelect: "none",
    };
    return (
        <Typography variant="caption" sx={sx}>
            Time:
        </Typography>
    );
}

const isAllowed100 = (values) => {
    const numValue = Number(values.value)
    if (isNaN(numValue)) return true
    return (numValue >= 0) && (numValue < 100)
}

const isAllowed60 = (values) => {
    const numValue = Number(values.value)
    if (isNaN(numValue)) return true
    return (numValue >= 0) && (numValue < 60)
}

const TimeHhFormat = (props: any) => {
    return (
        <NumberFormat
            decimalScale={0}
            allowNegative={false}
            isAllowed={isAllowed100}
            {...props}
        />
    );
};

const TimeMmFormat = (props: any) => {
    return (
        <NumberFormat
            decimalScale={0}
            allowNegative={false}
            isAllowed={isAllowed60}
            {...props}
        />
    );
};

const TimeSsFormat = (props: any) => {
    return (
        <NumberFormat
            decimalScale={3}
            fixedDecimalScale={true}
            allowNegative={false}
            isAllowed={isAllowed60}
            {...props}
        />
    );
};

export default function TimeJudgeDialog(props: TimeJudgeDialogProps) {
    const { open, bib, tag, startOrFinish, timeJudgeData, onChange, onClose, } = props;
    const { hh, mm, ss, judge } = timeJudgeData;

    const onChangeValues = (name: string, value: any) => {
        onChange({
            ...timeJudgeData,
            [name]: value,
        });
    };

    const onChangeHh = (e: any) => {
        onChangeValues('hh', e.target.value);
    };

    const onChangeMm = (e: any) => {
        onChangeValues('mm', e.target.value);
    };

    const onChangeSs = (e: any) => {
        onChangeValues('ss', e.target.value);
    };

    const onSelectedJudge = (judge: string) => {
        onChangeValues('judge', judge);
    };

    const onCancel = () => {
        onClose(false);
    };

    const onApply = () => {
        onClose(true);
    };

    const onFocus = (e: any) => {
        e.target.select();
    }

    const timeSx = {
        width: "40px",
        input: {
            height: "40px",
            textAlign: "right",
            fontSize: "1.6rem"
        }
    };

    const timeSsSx = {
        ...timeSx,
        width: "90px"
    };

    return (
        <Dialog open={open} onClose={onCancel}>
            <DialogTitle>
                {startOrFinish}: [{bib}] {tag}
            </DialogTitle>
            <DialogContent>
                <Stack spacing={1} direction="column">
                    <Stack spacing={1} direction="row">
                        <TimeCaption />
                        <TextField
                            value={hh}
                            onChange={onChangeHh}
                            label="hh"
                            variant="standard"
                            InputLabelProps={{ shrink: true }}
                            sx={timeSx}
                            InputProps={{
                                inputComponent: TimeHhFormat
                            }}
                            onFocus={onFocus}
                        />
                        <TextField
                            value={mm}
                            onChange={onChangeMm}
                            label="mm"
                            variant="standard"
                            InputLabelProps={{ shrink: true }}
                            sx={timeSx}
                            InputProps={{
                                inputComponent: TimeMmFormat
                            }}
                            onFocus={onFocus}
                        />
                        <TextField
                            value={ss}
                            onChange={onChangeSs}
                            label="ss.SS"
                            variant="standard"
                            InputLabelProps={{ shrink: true }}
                            sx={timeSsSx}
                            InputProps={{
                                inputComponent: TimeSsFormat
                            }}
                            onFocus={onFocus}
                        />
                    </Stack>
                    <JudgeButton
                        startOrFinish={startOrFinish}
                        judge={judge}
                        onSelected={onSelectedJudge}
                    />
                </Stack>
            </DialogContent>
            <DialogActions>
                <Button onClick={onCancel}>Cancel</Button>
                <Button onClick={onApply}>OK</Button>
            </DialogActions>
        </Dialog>
    );
}
