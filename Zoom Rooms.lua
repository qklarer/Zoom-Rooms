--#region Tables and Variables and Such

parseDebug = false
json = require('json')
Initializing = true
getParticipants = false
Error = false
isCall = false
inMeeting = false
inMeetingState = false

micMuteState = nil
videoInputSelectState = nil
audioInputSelectState = nil
audioOutputSelectState = nil
videoTableLength = nil
audioInputTableLength = nil
audioOutputTableLength = nil
bookingSelectValue = nil
participantSelectValue = nil
callJID = nil
waitingFor = nil
getStatus = nil

initializeCount = 0
errorCounter = 0
statusTimer = 0
parseCount = 0
refreshTimer = 0
waitCounter = 0
bookingListSize = 0
participantListSize = 0
getParticipantsTimer = 0
bookingTimer = 230
testCount = 0
bufferMin = 10000
bufferMax = 0


function Clear()

    Initializing = true
    getStatus = nil
    Error = false
    micMuteState = nil
    videoInputSelectState = nil
    audioInputSelectState = nil
    audioOutputSelectState = nil
    videoTableLength = nil
    audioInputTableLength = nil
    audioOutputTableLength = nil
    bookingSelectValue = nil
    initializeCount = 0
    errorCounter = 0
    statusTimer = 0
    parseCount = 0
    Connected = false
    NamedControl.SetValue("videoInputSelect", 1)
    NamedControl.SetValue("audioInputSelect", 1)
    NamedControl.SetValue("audioOutputSelect", 1)
    NamedControl.SetPosition("Connected", 0)
    NamedControl.SetText("State", "Off Line")
    NamedControl.SetText("meetingStatus", "")
    NamedControl.SetText("callerID", "")
    NamedControl.SetText("callStatus", "")
    for i = 1, 8 do
        NamedControl.SetText("Meeting" .. i, "")
        NamedControl.SetText("Participant" .. i, "")
        NamedControl.SetPosition("muteLED" .. i, 0)
        NamedControl.SetPosition("videoMuteLED" .. i, 0)
    end
    NamedControl.SetText("roomName", "")
    NamedControl.SetText("meetingNumber", "")
    NamedControl.SetText("currentTime", "")
    NamedControl.SetText("selectedBookName", "")
    NamedControl.SetText("selectedBookTime", "")
    NamedControl.SetText("selectedVideo", "")
    NamedControl.SetText("selectedAudioInput", "")
    NamedControl.SetText("selectedAudioOutput", "")
    NamedControl.SetText("videoInput", "")
    NamedControl.SetText("audioInput", "")
    NamedControl.SetText("audioOutput", "")
    SSH:Disconnect()
    --System.ClearDebugging()
    audioInputSources = {}
    audioOutputSources = {}
    cameraSources = {}
    bookingList = { startTime = {}, endTime = {}, calendarID = {}, isPrivate = {}, meetingName = {}, startEndTime = {},
        meetingNumber = {} }
    participantsList = { userID = {}, userName = {}, audioState = {}, videoState = {} }
end

-- Commands as that Zoom expects
zCommands =
{
    { Cmd = "zCommand Call ListParticipants " },
    { Cmd = "zCommand Phonebook List Offset: 0 Limit: 3000 " },
    { Cmd = "zCommand Bookings List " },
    { Cmd = "zCommand Call Accept callerJID: " },
    { Cmd = "zCommand Call Reject callerJID: " },
    { Cmd = "zCommand Call Info" },
    { Cmd = "zCommand Call Disconnect " },
    { Cmd = "zCommand Call Leave " },

    { Cmd = "zCommand Call HostChange Id: " },
    { Cmd = "zCommand Call Invite " },
    { Cmd = "zCommand Dial Start meetingNumber: " },
    { Cmd = "zCommand Dial Join meetingNumber: " },
    { Cmd = "zCommand Dial StartPmi Duration: " },
    { Cmd = "zCommand Schedule Delete MeetingNumber: " },
    { Cmd = "zcommand call listparticipants" },

    { Cmd = "zCommand Call Sharing HDMI ", State = "" },
    { Cmd = "zCommand Call ShareCamera id: ", State = "" },
    { Cmd = "zCommand Call MuteAll mute: ", State = "" },
    { Cmd = "zCommand Call MuteParticipant mute: ", State = NamedControl.GetPosition("muteParticipant") },
    { Cmd = "zCommand Call MuteParticipantVideo Mute: ", State = NamedControl.GetPosition("muteParticipantVideo") },


}

-- Control Names for buttons
zCommandNames = {
    [1] = "listParticipants",
    [2] = "listContacts",
    [3] = "listBookings",
    [4] = "acceptCall",
    [5] = "rejectCall",
    [6] = "callInfo", -- If in a meeting will show current meeting info. If no in meeting it will return "Not in meeting"
    [7] = "callDisconnect",
    [8] = "Leave",
    [9] = "hostChange",
    [10] = "Invite",
    [11] = "startMeeting",
    [12] = "joinMeeting",
    [13] = "startPmiMeeting",
    [14] = "deleteBooking",
    [15] = "listParticipants",

    [16] = "shareHDMI",
    [17] = "shareCamera",
    [18] = "muteAll",
    [19] = "muteParticipant",
    [20] = "muteParticipantVideo",

}

zConfiguration = {
    [1] = { Cmd = "zConfiguration Call Sharing optimize_video_sharing: ", State = "" },
    [2] = { Cmd = "zConfiguration Call Microphone Mute: ", State = "" },
    [3] = { Cmd = "zConfiguration Call Camera Mute: ", State = "" },
    [4] = { Cmd = "zConfiguration Audio Input is_sap_disabled: ", State = "" },
    [5] = { Cmd = "zConfiguration Audio Input reduce_reverb: ", State = "" },
    [6] = { Cmd = "zConfiguration Audio Input volume: ", State = "" },
    [7] = { Cmd = "zConfiguration Audio Output volume: ", State = "" },
    [8] = { Cmd = "zConfiguration Video hide_conf_self_video: ", State = "" },
    [9] = { Cmd = "zConfiguration Video Camera Mirror: ", State = "" },
    [10] = { Cmd = "zConfiguration Call Layout ShareThumb: ", State = "" },
    [11] = { Cmd = "zConfiguration Call Layout Style: ", State = "" },
    [12] = { Cmd = "zConfiguration Call Layout Size: ", State = "" },
    [13] = { Cmd = "zConfiguration Call Layout Position: ", State = "" },
    [14] = { Cmd = "zConfiguration Call Lock Enable: ", State = "" },
    [15] = { Cmd = "zConfiguration Call ClosedCaption Visible: ", State = "" },
    [16] = { Cmd = "zConfiguration Call ClosedCaption FontSize: ", State = "" },
    [17] = { Cmd = "zConfiguration SIPCall Microphone Mute: ", State = "" },

    [18] = { Cmd = "zConfiguration Audio Input selectedId: ", State = "" },
    [19] = { Cmd = "zConfiguration Audio Output selectedId: ", State = "" },
    [20] = { Cmd = "zConfiguration Video Camera selectedId: ", State = "" },
}

zConfigurationNames = {
    [1] = "callSharingOptimizeVideoSharing",
    [2] = "callMicrophoneMute",
    [3] = "callCameraMute",
    [4] = "audioInput_SAP_Disabled",
    [5] = "audioInputReduce_Reverb",
    [6] = "audioInputVolume",
    [7] = "audioOutputVolume",
    [8] = "videoConfSelfVideo",
    [9] = "videoCameraMirror",
    [10] = "callLayoutShareThumb",
    [11] = "callLayoutStyle",
    [12] = "callLayoutSize",
    [13] = "callLayoutPosition",
    [14] = "callLockEnable",
    [15] = "callClosedCaptionVisible",
    [16] = "callClosedCaptionFontSize",
    [17] = "sipCallMicrophoneMute",

    [18] = "audioInputSelected",
    [19] = "audioOutputSelected",
    [20] = "videoCameraSelected",
}

zStatus = {
    [1]  = { Cmd = "zstatus call status" },
    [2]  = { Cmd = "zStatus Audio Input Line" },
    [3]  = { Cmd = "zStatus Audio Output Line" },
    [4]  = { Cmd = "zStatus Video Camera Line" },
    [5]  = { Cmd = "zStatus Video Optimizable" },
    [6]  = { Cmd = "zStatus SystemUnit" },
    [7]  = { Cmd = "zStatus Capabilities" },
    [8]  = { Cmd = "zStatus Sharing" },
    [9]  = { Cmd = "zStatus CameraShare" },
    [10] = { Cmd = "zStatus Call Layout" },
    [11] = { Cmd = "zStatus Call ClosedCaption Available" },
    [12] = { Cmd = "zstatus NumberOfScreens" }
}

zStatusNames = {
    [1] = "call",
    [2] = "audioInput",
    [3] = "audioOutput",
    [4] = "cameraLine",
    [5] = "videoOptimizable",
    [6] = "systemUnit",
    [7] = "Capabilities",
    [8] = "Sharing",
    [9] = "cameraShare",
    [10] = "callLayout",
    [11] = "closedCaptionAvailable",
    [12] = "numberOfScreens"
}

boolCommands = {
    "callMicrophoneMute", "callCameraMute", "audioInputReduce_Reverb",
    "callSharingOptimizeVideoSharing", "audioInput_SAP_Disabled", "videoConfSelfVideo", "videoCameraMirror", "shareHDMI",
    "shareCamera", "callClosedCaptionVisible"
}

intCommands = {
    "audioInputVolume", "audioOutputVolume"
}

startStopCommands = {
    "shareHDMI", "shareCamera"
}

audioInputSources = {}
audioOutputSources = {}
cameraSources = {}
bookingList = { startTime = {}, endTime = {}, calendarID = {}, isPrivate = {}, meetingName = {}, startEndTime = {},
    meetingNumber = {} }
participantsList = { userID = {}, userName = {}, audioState = {} }

SSH = Ssh.New()
SSH.ReadTimeout = 15
SSH.WriteTimeout = 15
SSH.ReconnectTimeout = 10

--#endregion Tables and Variables and Such

--#region Helpers and Other Functions

function Wait(Cmd)
    zStatusSend(Cmd)
    -- zStatusSend(zStatus[2])
    -- zStatusSend(zStatus[3])
    -- zStatusSend(zStatus[4])
    --waitingFor = nil
end

function Format(str)

    return (string.gsub(str, ".0", ""))
end

function Split(s, delimiter)

    local result = {}

    for match in (s .. delimiter):gmatch("(.-)" .. delimiter) do
        table.insert(result, match)
    end
end

function FormarCurrentTime()

    local currentTimeHour = tonumber(os.date("%H"))
    local currentTimeMinute = tonumber(os.date("%M"))
    local Period
    if currentTimeHour >= 13 then
        Period = "PM"
        currentTimeHour = currentTimeHour - 12
    else
        Period = "AM"
    end

    NamedControl.SetText("currentTime", currentTimeHour .. ":" .. currentTimeMinute .. Period)

end

-- When scheduling using the Zoom Outlook extension, time is formatted differently (UTC)
function CheckUTC(startTime, endTime)

    local StartHour
    local StartMin
    local EndHour
    local EndMin
    local MeetTime
    -- If meeting time is in UTC, it's denoted with a 'Z' at the end for Zulu
    if startTime:find("Z") then
        StartHour, StartMin = string.match(AdjustUTC(startTime), "(%d+):(%d+)")
        EndHour, EndMin = string.match(AdjustUTC(endTime), "(%d+):(%d+)")
    else
        StartHour, StartMin = string.match(startTime, "(%d+):(%d+)")
        EndHour, EndMin = string.match(endTime, "(%d+):(%d+)")
    end
    MeetTime = AdjustTimeFormat(StartHour, StartMin, EndHour, EndMin)

    return MeetTime
end

function AdjustUTC(time)

    --  year  month  day  hour   min   sec
    local y, mon, d, h, m, s = time:match("(%d+)-(%d+)-(%d+)T(%d+):(%d+):(%d+)Z")
    local Time =
    {
        year  = tonumber(y),
        month = tonumber(mon),
        day   = tonumber(d),
        hour  = tonumber(h),
        min   = tonumber(m),
        sec   = tonumber(s)
    }

    -- that time is in UTC but we need to convert to local time before creating the time object
    -- '*t' returns date in local time zone
    -- '!*t' returns date in utc time
    -- time gets seconds since epoch and by subracting them we get the difference in seconds
    local localUTCOffsetInSeconds = os.time(os.date('*t')) - os.time(os.date('!*t'))
    --print("localUTCOffsetInSeconds", localUTCOffsetInSeconds)
    -- Check for Daylight Savings Time
    for k, v in pairs(os.date('*t')) do
        -- Adjust hour if DST is true (it's the only boolean returned by os.date)
        if v == true then Time.hour = Time.hour + 1 end
    end
    -- Add in our offset
    Time = os.time(Time) + localUTCOffsetInSeconds
    temp = os.date("*t", Time)
    currentTime = temp.hour .. ":" .. temp.min

    return currentTime
end

-- Format to readable 12-hr time
function AdjustTimeFormat(StartHour, StartMin, EndHour, EndMin)
    local time

    if tonumber(StartHour) >= 12 and tonumber(EndHour) >= 12 then
        -- When meeting times are all PM adjust from UTC to standard by subtracting 12
        if StartHour ~= "12" then StartHour = StartHour - 12 end
        if EndHour ~= "12" then EndHour = EndHour - 12 end
        time = string.format("%i:%.2i PM - %i:%.2i PM", StartHour, StartMin, EndHour, EndMin)
    elseif tonumber(StartHour) > 12 then
        -- Meet begins in the PM and concludes in the AM
        StartHour = StartHour - 12
        if EndHour == "00" then EndHour = "12" end
        time = string.format("%i:%.2i PM - %i:%.2i AM", StartHour, StartMin, EndHour, EndMin)
    elseif tonumber(EndHour) > 12 then
        -- Meeting begins in the AM and concludes in the PM
        EndHour = EndHour - 12
        if StartHour == "00" then StartHour = "12" end
        time = string.format("%i:%.2i AM - %i:%.2i PM", StartHour, StartMin, EndHour, EndMin)
    else
        -- Entirety of meeting takes places in the AM
        if StartHour == "00" then StartHour = "12" end
        if EndHour == "00" then EndHour = "12" end
        time = string.format("%i:%.2i AM - %i:%.2i AM", StartHour, StartMin, EndHour, EndMin)
    end
    return time
end

-- Gathers errors from multiple places
function IsError(e)
    Error = true
    Initializing = true
    Connected = false
    NamedControl.SetText("Error", e)
end

-- Makes a booking
function Book()

    ParseResponse()
    if NamedControl.GetPosition("newBookingPrivate") == 1 then
        Private = "on"
    elseif NamedControl.GetPosition("newBookingPrivate") == 0 then
        Private = "off"
    end

    SSH:Write("zCommand Schedule Add MeetingName: " ..
        NamedControl.GetText("newBookingName") ..
        " Start: " ..
        NamedControl.GetText("newBookingStart") .. -- YYYYMMDDHHMM
        " End: " .. NamedControl.GetText("newBookingEnd") .. " private: " .. Private .. "\r")
    parseTimer = true
    zCommandSend("", zCommands[3].Cmd)
end

function SetParticipants()

    local Par = NamedControl.GetValue("participantSelect")
    NamedControl.SetText("selectedParticipantName",
        participantsList.userName[Par])

    NamedControl.SetPosition("muteParticipant", participantsList.audioState[Par])
    NamedControl.SetPosition("muteParticipantVideo", participantsList.videoState[Par])
    participantSelectValue = Par

end

--#endregion Helpers and Other Functions

--#region Parsers

function ParseResponse() -- function that reads the SSH TCP socket

    -- On some occations, the rx buffer will contain largeamounts of information, usally regarding
    -- different countrys meeting codes. This information is not helpful and is difficult to parse
    -- When this happens we can just return from the function to avoid errors
    if SSH.BufferLength > bufferMax then bufferMax = SSH.BufferLength
    elseif SSH.BufferLength < bufferMin and SSH.BufferLength > 0 then bufferMin = SSH.BufferLength
    end

    if SSH.BufferLength < 50 or SSH.BufferLength > 10000 then
        SSH:Read(SSH.BufferLength)
        bufferMin = 10000
        bufferMax = 0
        return
    end
    rx = SSH:Read(SSH.BufferLength) -- assign the contents of the buffer to a variable

    if rx ~= nil then
        if rx:match("*e Connection rejected") then
            if parseDebug then print("Debug: ", 1) end
            IsError("Unknown Error")
            return
        elseif rx:match("zStatus Audio Input Line") then
            if parseDebug then
                print(rx)
                print("Debug: ", 2)
            end
            audioInRX = rx:gsub("zStatus Audio Input Line", "")
            local jsonStart = (string.find(audioInRX, "{"))
            local decodedJson = json.decode(audioInRX, jsonStart)
            if decodedJson["Audio Input Line"] ~= nil then
                audioInputTableLength = #decodedJson["Audio Input Line"]
                for i = 1, audioInputTableLength do
                    audioInputSources[i] = decodedJson["Audio Input Line"][i]
                end
                for k, v in pairs(audioInputSources) do
                    if v.Selected then
                        NamedControl.SetText("selectedAudioInput", v.Name)
                    end
                end
            end
            return
        elseif rx:match("zStatus Audio Output Line") then
            if parseDebug then
                print(rx)
                print("Debug: ", 3)
            end
            audioOutRX = rx:gsub("zStatus Audio Output Line", "")
            local jsonStart = (string.find(audioOutRX, "{"))
            local decodedJson = json.decode(audioOutRX, jsonStart)
            if decodedJson["Audio Output Line"] ~= nil then
                audioOutputTableLength = #decodedJson["Audio Output Line"]

                for i = 1, audioOutputTableLength do
                    audioOutputSources[i] = decodedJson["Audio Output Line"][i]
                end

                for k, v in pairs(audioOutputSources) do
                    if v.Selected then
                        NamedControl.SetText("selectedAudioOutput", v.Name)
                    end
                end
            end
            return
        elseif rx:match("zStatus Video Camera Line") then

            if parseDebug then
                print(rx)
                print("Debug: ", 4)
            end
            camRX = rx:gsub("zStatus Video Camera Line", "")
            local jsonStart = (string.find(camRX, "{"))
            local decodedJson = json.decode(camRX, jsonStart)
            if decodedJson["Video Camera Line"] ~= nil then
                videoTableLength = #decodedJson["Video Camera Line"]
                for i = 1, videoTableLength do
                    cameraSources[i] = decodedJson["Video Camera Line"][i]
                end
                for k, v in pairs(cameraSources) do
                    if v.Selected then
                        NamedControl.SetText("selectedVideo", v.Name)
                    end
                end
            end
            return
        elseif rx:match("zConfiguration Video Camera selectedId: ") then
            if parseDebug then print("Debug: ", 5) end
            getStatus = "cameraLine"
            return
        elseif rx:match("zConfiguration Audio Input selectedId: ") then
            if parseDebug then print("Debug: ", 6) end
            getStatus = "audioInput"
            return
        elseif rx:match("zConfiguration Audio Output selectedId: ") then
            if parseDebug then print("Debug: ", 7) end
            getStatus = "audioOutput"
            return
        elseif rx:match("ERROR") then
            if parseDebug then print("Debug: ", 8) end
            NamedControl.SetText("Error", "Error")
            Error = true
            return
        elseif rx:match("zStatus SystemUnit") then
            local meetingNumber
            if parseDebug then print("Debug: ", 9) end
            local systemUnitRX = rx:gsub("zStatus SystemUnit", "")
            local decodedJson = json.decode(systemUnitRX)
            local State = decodedJson.Status.state
            if decodedJson.SystemUnit ~= nil then
                meetingNumber = decodedJson.SystemUnit.meeting_number
            end
            local roomName = decodedJson.SystemUnit.room_info.room_name
            NamedControl.SetText("meetingNumber", meetingNumber)
            NamedControl.SetText("roomName", roomName)
            NamedControl.SetText("State", State)
            return
        elseif rx:match("zCommand Bookings List") then
            if parseDebug then print("Debug: ", 10) end
            ParseCalander(rx)
            return
        elseif rx:match("zcommand call listparticipants") then
            if parseDebug then print("Debug: ", 11) end
            ParseParticipants(rx)
            return
        elseif rx:match("zstatus call status") then
            local Status
            if parseDebug then print("Debug: ", 12) end
            local callStatusRX = rx
            callStatusRX = callStatusRX:gsub("zstatus call status", "")
            local jsonStart = string.find(callStatusRX, "{")


            local decodedJson = json.decode(callStatusRX, jsonStart)

            if decodedJson.Call ~= nil then
                Status = decodedJson.Call.Status
            end
            if Status == "NOT_IN_MEETING" then
                inMeeting = false
                NamedControl.SetText("meetingStatus", "")

            elseif Status == "IN_MEETING" then
                if parseDebug then print("Debug: ", 13) end
                inMeeting = true
                NamedControl.SetText("meetingStatus", "In Meeting")
                NamedControl.SetText("callStatus", "")
            end
            if rx:match("IncomingCallIndication") then

                -- Rejected or ignored calls are reported as "TreatedIncomingCallIndication"
                -- check if the call is ignored first, if it was answered check again to see
                -- the information regarding the accepted call.
                jsonStart1 = string.find(callStatusRX, "Treated")

                if jsonStart1 == nil then
                    jsonStart1 = string.find(callStatusRX, "IncomingCallIndication")
                    t = json.decode(callStatusRX, jsonStart1 - 9)
                    callJID = t.IncomingCallIndication.callerJID
                    local callerName = t.IncomingCallIndication.callerName
                    local callStatus = "Incoming Call"
                    NamedControl.SetText("callStatus", callStatus)
                    NamedControl.SetText("callerID", callerName)
                    isCall = true
                else
                    t = json.decode(callStatusRX, jsonStart1 - 85)
                    if t.TreatedIncomingCallIndication.accepted == false then
                        NamedControl.SetText("callStatus", "")
                        NamedControl.SetText("callerID", "")
                        isCall = false
                    end
                end
            end
            return
            -- IncomingCallIndication can be included under a call status
            -- update in which case it is treated as a sub catch of call status,
            -- or as its own parsed response as seen bellow.
        elseif rx:match("IncomingCallIndication") then

            if parseDebug then print("Debug: ", 14) end
            local decodedJson = json.decode(rx)
            if decodedJson.IncomingCallIndication ~= nil then
                callJID = decodedJson.IncomingCallIndication.callerJID
                local callerName = decodedJson.IncomingCallIndication.callerName
                local callStatus = "Incoming Call"
                NamedControl.SetText("callStatus", callStatus)
                NamedControl.SetText("callerID", callerName)
                isCall = true
            elseif decodedJson.TreatedIncomingCallIndication ~= nil then
                if decodedJson.TreatedIncomingCallIndication.accepted == false then
                    NamedControl.SetText("callStatus", "")
                    NamedControl.SetText("callerID", "")
                    isCall = false
                end
            end
            return
        elseif rx:match("CallAcceptRejectResult") then

            NamedControl.SetText("callStatus", "")
            NamedControl.SetText("callerID", "")
            return
        elseif rx:match("TreatedIncomingCallIndication") then

            local decodedJson = json.decode(rx)
            if decodedJson.TreatedIncomingCallIndication.accepted == false then
                NamedControl.SetText("callStatus", "")
                NamedControl.SetText("callerID", "")
                return
            end
        end
    end
end

function ParseCalander(Calendar)

    bookingList = { startTime = {}, endTime = {}, calendarID = {}, isPrivate = {}, meetingName = {},
        startEndTime = {},
        meetingNumber = {} }
    calendarRX = Calendar:gsub("zCommand Bookings List", "")

    local jsonStart = (string.find(calendarRX, "BookingsListResult"))
    if jsonStart == nil then
        return
    end

    local decodedJson = json.decode(calendarRX, jsonStart - 9) -- 9 characters untill start of usable json object

    if decodedJson.BookingsListResult ~= nil then
        bookingListSize = #decodedJson.BookingsListResult

        if bookingListSize ~= nil then
            for i = 1, bookingListSize do
                table.insert(bookingList.startTime, decodedJson.BookingsListResult[i].startTime)
                table.insert(bookingList.endTime, decodedJson.BookingsListResult[i].endTime)
                table.insert(bookingList.calendarID, decodedJson.BookingsListResult[i].calendarID)
                table.insert(bookingList.isPrivate, decodedJson.BookingsListResult[i].isPrivate)
                table.insert(bookingList.meetingName, decodedJson.BookingsListResult[i].meetingName)
                table.insert(bookingList.meetingNumber, decodedJson.BookingsListResult[i].meetingNumber)
                table.insert(bookingList.startEndTime, CheckUTC(bookingList.startTime[i], bookingList.endTime[i]))

            end

            if bookingList.meetingName[1] == nil then
                NamedControl.SetText("Meeting1", "")
            elseif bookingList.meetingName ~= nil then
                NamedControl.SetText("Meeting1", bookingList.meetingName[1] .. " " .. bookingList.startEndTime[1])
            end

            if bookingList.meetingName[2] == nil then
                NamedControl.SetText("Meeting2", "")
            elseif bookingList.meetingName ~= nil then
                NamedControl.SetText("Meeting2", bookingList.meetingName[2] .. " " .. bookingList.startEndTime[2])
            end

            if bookingList.meetingName[3] == nil then
                NamedControl.SetText("Meeting3", "")
            elseif bookingList.meetingName ~= nil then
                NamedControl.SetText("Meeting3", bookingList.meetingName[3] .. " " .. bookingList.startEndTime[3])
            end

            if bookingList.meetingName[4] == nil then
                NamedControl.SetText("Meeting4", "")
            elseif bookingList.meetingName ~= nil then
                NamedControl.SetText("Meeting4", bookingList.meetingName[4] .. " " .. bookingList.startEndTime[4])
            end

            if bookingList.meetingName[5] == nil then
                NamedControl.SetText("Meeting5", "")
            elseif bookingList.meetingName ~= nil then
                NamedControl.SetText("Meeting5", bookingList.meetingName[5] .. " " .. bookingList.startEndTime[5])
            end

            if bookingList.meetingName[6] == nil then
                NamedControl.SetText("Meeting6", "")
            elseif bookingList.meetingName ~= nil then
                NamedControl.SetText("Meeting6", bookingList.meetingName[6] .. " " .. bookingList.startEndTime[6])
            end

            if bookingList.meetingName[7] == nil then
                NamedControl.SetText("Meeting7", "")
            elseif bookingList.meetingName ~= nil then
                NamedControl.SetText("Meeting7", bookingList.meetingName[7] .. " " .. bookingList.startEndTime[7])
            end

            if bookingList.meetingName[8] == nil then
                NamedControl.SetText("Meeting8", "")
            elseif bookingList.meetingName ~= nil then
                NamedControl.SetText("Meeting8", bookingList.meetingName[8] .. " " .. bookingList.startEndTime[8])
            end
            NamedControl.SetText("selectedBookName", bookingList.meetingName[1])
            NamedControl.SetText("selectedBookTime", bookingList.startEndTime[1])
        end
    end
end

function ParseParticipants(Participants)

    local audioValue
    local videoValue
    local decodedJson
    participantsList = { userID = {}, userName = {}, audioState = {}, videoState = {} }
    participantsRX = Participants:gsub("zcommand call listparticipants", "")

    local jsonStart = (string.find(participantsRX, "ListParticipantsResult"))
    if jsonStart ~= nil then
        decodedJson = json.decode(participantsRX, jsonStart - 9) -- 9 characters untill start of usable json object

        if decodedJson.ListParticipantsResult ~= nil then
            participantListSize = #decodedJson.ListParticipantsResult
            for i = 1, participantListSize do
                table.insert(participantsList.userName, decodedJson.ListParticipantsResult[i].user_name)
                table.insert(participantsList.userID, decodedJson.ListParticipantsResult[i].user_id)
                for k, v in pairs(decodedJson.ListParticipantsResult[i]) do
                    if k == "audio_status state" then
                        if v == "AUDIO_MUTED" then
                            audioValue = 1
                        elseif v == "AUDIO_UNMUTED" then
                            audioValue = 0
                        end
                        table.insert(participantsList.audioState, audioValue)
                        NamedControl.SetPosition("muteLED" .. i, participantsList.audioState[i])
                        -- NamedControl.SetPosition("callMicrophoneMute", participantsList.audioState[1])
                    end
                    if k == "video_status is_sending" then
                        if v == true then
                            videoValue = 0
                        elseif v == false then
                            videoValue = 1
                        end
                        table.insert(participantsList.videoState, videoValue)
                        NamedControl.SetPosition("videoMuteLED" .. i, participantsList.videoState[i])
                        --  NamedControl.SetPosition("callCameraMute", participantsList.videoState[1])
                    end
                end
            end
        end
    end

    if participantsList.userName[1] == nil then
        NamedControl.SetText("Participant1", "")
        NamedControl.SetPosition("videoMuteLED1", 0)
        NamedControl.SetPosition("muteLED1", 0)
    elseif participantsList.userName ~= nil then
        NamedControl.SetText("Participant1", participantsList.userName[1])
        NamedControl.SetText("selectedParticipantName", participantsList.userName[1])
        NamedControl.SetValue("participantSelect", 1)
        participantSelectValue = 0
    end
    if participantsList.userName[2] == nil then
        NamedControl.SetText("Participant2", "")
        NamedControl.SetPosition("videoMuteLED2", 0)
        NamedControl.SetPosition("muteLED2", 0)
    elseif participantsList.userName[2] ~= nil then
        NamedControl.SetText("Participant2", participantsList.userName[2])
    end
    if participantsList.userName[3] == nil then
        NamedControl.SetText("Participant3", "")
        NamedControl.SetPosition("videoMuteLED3", 0)
        NamedControl.SetPosition("muteLED3", 0)
    elseif participantsList.userName[3] ~= nil then
        NamedControl.SetText("Participant3", participantsList.userName[3])
    end
    if participantsList.userName[4] == nil then
        NamedControl.SetText("Participant4", "")
        NamedControl.SetPosition("videoMuteLED4", 0)
        NamedControl.SetPosition("muteLED4", 0)
    elseif participantsList.userName[4] ~= nil then
        NamedControl.SetText("Participant4", participantsList.userName[4])
    end
    if participantsList.userName[5] == nil then
        NamedControl.SetText("Participant5", "")
        NamedControl.SetPosition("videoMuteLED5", 0)
        NamedControl.SetPosition("muteLED5", 0)
    elseif participantsList.userName[5] ~= nil then
        NamedControl.SetText("Participant5", participantsList.userName[5])
    end
    if participantsList.userName[6] == nil then
        NamedControl.SetText("Participant6", "")
        NamedControl.SetPosition("videoMuteLED6", 0)
        NamedControl.SetPosition("muteLED6", 0)
    elseif participantsList.userName[6] ~= nil then
        NamedControl.SetText("Participant6", participantsList.userName[6])
    end
    if participantsList.userName[7] == nil then
        NamedControl.SetText("Participant7", "")
        NamedControl.SetPosition("videoMuteLED7", 0)
        NamedControl.SetPosition("muteLED7", 0)
    elseif participantsList.userName[7] ~= nil then
        NamedControl.SetText("Participant7", participantsList.userName[7])
    end
    if participantsList.userName[8] == nil then
        NamedControl.SetText("Participant8", "")
        NamedControl.SetPosition("videoMuteLED8", 0)
        NamedControl.SetPosition("muteLED8", 0)
    elseif participantsList.userName[8] ~= nil then
        NamedControl.SetText("Participant8", participantsList.userName[8])
    end
end

--#endregion Parsers

--#region SSH Callback
SSH.Connected = function() -- function called when the TCP socket is connected
    print("Socket connected")
    SSH:Write("format json \r\n")
    IsError("Socket connected")
    Connected = true
end

SSH.Reconnect = function() -- function called when the TCP socket is reconnected
    print("Socket reconnecting...")
    IsError("Socket reconnecting...")
end

SSH.Closed = function() -- function called when the TCP socket is closed
    print("Socket closed")
    IsError("Socket closed")
end

SSH.Error = function() -- function called when the TCP socket has an error
    print("Socket error")
    IsError("Socket error")
end

SSH.Timeout = function() -- function called when the TCP socket times out
    print("Socket timeout")
    IsError("Socket timeout")
end

SSH.LoginFailed = function() -- function called when SSH login fails
    print("SSH login failed")
    IsError("SSH login failed")
end

--SSH.Data = ParseResponse -- ParseResponse is called when the SSH object has data
--#endregion SSH Callback

--#region Send Commands
function zCommandSend(Name, Cmd, Value)
    --print(Name, Cmd, Value)
    if Value == nil then
        if Name == "startPmiMeeting" then
            SSH:Write(Cmd .. Format(NamedControl.GetValue("pmiDuration")) .. "\r\n")
            getParticipants = true
        elseif Name == "startMeeting" or Name == "deleteBooking" then
            if bookingList.meetingNumber[NamedControl.GetValue("bookingSelect")] ~= nil and Cmd ~= nil then
                SSH:Write(Cmd .. bookingList.meetingNumber[NamedControl.GetValue("bookingSelect")] .. "\r\n")
                zCommandSend("", zCommands[3].Cmd)
                getParticipants = true
            end
        elseif Name == "acceptCall" or Name == "rejectCall" then

            if callJID == nil then
                return
            else
                SSH:Write(Cmd .. callJID .. "\r\n")

            end

        else

            SSH:Write(Cmd .. "\r\n")
        end
    end

    if Value ~= nil then
        if Name == "shareHDMI" then
            if Value == 0 then
                Value = "Stop"
            elseif Value == 1 then
                Value = "Start"
            end
        elseif Name == "muteAll" then
            if Value == 0 then
                Value = "off"
            elseif Value == 1 then
                Value = "on"
            end
        elseif Name == "shareCamera" then
            if cameraSources ~= nil then
                if cameraSources[NamedControl.GetValue("videoInputSelect")].id ~= nil then
                    if Value == 0 then
                        Value = cameraSources[NamedControl.GetValue("videoInputSelect")].id .. " Status: off"
                    elseif Value == 1 then
                        Value = cameraSources[NamedControl.GetValue("videoInputSelect")].id .. " Status: on"
                    end
                end
            end
            SSH:Write(Cmd .. Value .. "\r\n")
        end
    end
    ParseResponse()
    parseTimer = true
end

function zConfigurationSend(Name, Cmd, Value)

    if Value ~= nil then

        for k, v in pairs(boolCommands) do
            if Name == v then
                if Value == 0 then
                    Value = "off"
                elseif Value == 1 then
                    Value = "on"
                end
            end
        end

        for k, v in pairs(intCommands) do
            if Name == v then
                Value = math.floor(Value)
            end
        end
        SSH:Write(Cmd .. " " .. Value .. "\r\n")
    elseif Value == nil then

        if Name == "videoCameraSelected" then
            SSH:Write(Cmd .. cameraSources[NamedControl.GetValue("videoInputSelect")].id .. "\r\n")
            waitingFor = zStatus[4].Cmd
        elseif Name == "audioInputSelected" then
            SSH:Write(Cmd .. audioInputSources[NamedControl.GetValue("audioInputSelect")].id .. "\r\n")
            waitingFor = zStatus[2].Cmd
        elseif Name == "audioOutputSelected" then
            SSH:Write(Cmd .. audioOutputSources[NamedControl.GetValue("audioOutputSelect")].id .. "\r\n")
            waitingFor = zStatus[3].Cmd
        end
    end
    ParseResponse()
    parseTimer = true
end

function zStatusSend(Cmd)
    SSH:Write(Cmd .. "\r\n")
    ParseResponse()
    parseTimer = true
end

--#endregion Send Commands

--#region Get Commands
function PollzCommands()
    for i = 1, 15 do
        if NamedControl.GetPosition(zCommandNames[i]) == 1 then
            zCommandSend(zCommandNames[i], zCommands[i].Cmd)
            NamedControl.SetPosition(zCommandNames[i], 0)
        end
    end

    for i = 16, 20 do
        if NamedControl.GetPosition(zCommandNames[i]) ~= zCommands[i].State then
            zCommands[i].State = NamedControl.GetPosition(zCommandNames[i])
            zCommandSend(zCommandNames[i], zCommands[i].Cmd, zCommands[i].State)
            -- print(zCommandNames[i], zCommands[i].Cmd, zCommands[i].State)
        end
    end
end

function PollzConfiguration()

    for i = 1, 17 do
        if NamedControl.GetValue(zConfigurationNames[i]) ~= zConfiguration[i].State then
            zConfiguration[i].State = NamedControl.GetValue(zConfigurationNames[i])
            zConfigurationSend(zConfigurationNames[i], zConfiguration[i].Cmd, zConfiguration[i].State)
            -- print(zConfigurationNames[i], zConfiguration[i].Cmd, zConfiguration[i].State)
        end
    end
    for i = 18, 20 do
        if NamedControl.GetPosition(zConfigurationNames[i]) == 1 then
            zConfigurationSend(zConfigurationNames[i], zConfiguration[i].Cmd)
            NamedControl.SetPosition(zConfigurationNames[i], 0)
        end
    end
end

function PollzStatus(Status)

    for k, v in pairs(zStatusNames) do
        if v == Status then
            zStatusSend(zStatus[k].Cmd)
        end
    end
end

--#endregion Get Commands

function Initialize()

    initializeCount = initializeCount + 1

    if initializeCount == 2 then
        NamedControl.SetText("State", "Connecting...")
        ParseResponse()
        zStatusSend(zStatus[2].Cmd) -- Audio Input
    elseif initializeCount == 4 then
        ParseResponse()
    elseif initializeCount == 6 then
        zStatusSend(zStatus[3].Cmd) -- Audio Output
    elseif initializeCount == 8 then
        ParseResponse()
    elseif initializeCount == 10 then
        zStatusSend(zStatus[4].Cmd) -- Video
    elseif initializeCount == 12 then
        ParseResponse()
    elseif initializeCount == 13 then
        zStatusSend(zStatus[6].Cmd) -- System Unit
    elseif initializeCount == 15 then
        ParseResponse()
    elseif initializeCount == 17 then
        zCommandSend("", zCommands[3].Cmd) -- Calendar
    elseif initializeCount == 19 then
        ParseResponse()
        FormarCurrentTime()
        --NamedControl.SetText("currentTime", (os.date("%H:%M")))
        Initializing = false
        initializeCount = 0
    end
end

function TimerClick()

    if NamedControl.GetPosition("Connect") == 1 then
        SSH:Connect(NamedControl.GetText("IP"), 2244, NamedControl.GetText("Username"), NamedControl.GetText("Password"))
        NamedControl.SetPosition("Connect", 0)
    end

    if NamedControl.GetPosition("Disconnect") == 1 then
        Clear()
        NamedControl.SetPosition("Disconnect", 0)
    end

    if Error then
        errorCounter = errorCounter + 1
        if errorCounter == 8 then
            errorCounter = 0
            Error = false
            NamedControl.SetText("Error", "")
        end
    end
    if NamedControl.GetPosition("Send") == 1 then
        ParseResponse()
        zStatusSend("zStatus Call ClosedCaption Available")
        NamedControl.SetPosition("Send", 0)
    end

    if Connected then

        if Initializing then
            Initialize()
        end

        if not Initializing then

            if SSH.IsConnected then
                NamedControl.SetPosition("Connected", 1)
            else
                NamedControl.SetPosition("Connected", 0)
            end

            PollzCommands()
            PollzConfiguration()

            --#region Timers

            refreshTimer = refreshTimer + 1
            if refreshTimer == 10 then
                ParseResponse()
                zStatusSend(zStatus[1].Cmd) -- Get Call Status
            elseif refreshTimer == 20 then
                ParseResponse()
                FormarCurrentTime()
                NamedControl.SetText("currentTime", (os.date("%H:%M")))

                refreshTimer = 0
            end

            bookingTimer = bookingTimer + 1
            if bookingTimer == 240 then
                zCommandSend("", zCommands[3].Cmd) -- Get Booking List
                bookingTimer = 0
            end

            if getStatus ~= nil then
                statusTimer = statusTimer + 1
                if statusTimer == 11 then
                    PollzStatus(getStatus) -- Get overall status
                    statusTimer = 0
                    getStatus = nil
                end
            end

            if parseTimer then
                parseCount = parseCount + 1
                if parseCount == 12 then
                    ParseResponse() -- Parse the rx buffer to clear it of any unwanted data
                    parseCount = 0
                    parseTimer = false
                end
            end
            if getParticipants or inMeeting then
                getParticipantsTimer = getParticipantsTimer + 1
                if getParticipantsTimer == 15 then
                    ParseResponse()
                    zCommandSend("", zCommands[15].Cmd) -- If in a meeting, check for participant in case they join or leave after start of meeting
                    getParticipantsTimer = 0
                    getParticipants = false
                end
            end

            if waitingFor ~= nil then -- When changing AV IO Zoom Room takes time to reflect the change, we need to wait to get a status update. As well, depending on if someone changes multiple devices quickly, we should check for IO regardless of what was changed when.
                waitCounter = waitCounter + 1
                if waitCounter == 8 then
                    Wait(zStatus[2].Cmd)
                elseif waitCounter == 16 then
                    Wait(zStatus[3].Cmd)
                elseif waitCounter == 24 then
                    Wait(zStatus[4].Cmd)
                    waitingFor = nil
                    waitCounter = 0

                end

            end

            --#endregion Timers

            if NamedControl.GetPosition("Book") == 1 then
                Book()
                NamedControl.SetPosition("Book", 0)
            end

            -- Buttons in GUI have a undetermined value until Zoom rooms defines that value
            -- Once the value is determined for a button, compare the state of that button
            -- to the actual total value in Zoom. Comparing the zoom value and button value
            -- keeps them in sync
            if NamedControl.GetValue("bookingSelect") > bookingListSize then
                NamedControl.SetValue("bookingSelect", bookingListSize)

            elseif NamedControl.GetValue("bookingSelect") ~= bookingSelectValue then

                NamedControl.SetText("selectedBookName", bookingList.meetingName[NamedControl.GetValue("bookingSelect")])
                NamedControl.SetText("selectedBookTime",
                    bookingList.startEndTime[NamedControl.GetValue("bookingSelect")
                    ])
                bookingSelectValue = NamedControl.GetValue("bookingSelect")
            end

            if audioOutputTableLength ~= nil and videoTableLength ~= nil and audioInputTableLength ~= nil then
                if NamedControl.GetValue("videoInputSelect") > videoTableLength then
                    NamedControl.SetValue("videoInputSelect", videoTableLength)

                elseif NamedControl.GetValue("videoInputSelect") ~= videoInputSelectState then
                    NamedControl.SetText("videoInput", cameraSources[NamedControl.GetValue("videoInputSelect")].Name)
                    videoInputSelectState = NamedControl.GetValue("videoInputSelect")
                end

                if NamedControl.GetValue("audioInputSelect") > audioInputTableLength then
                    NamedControl.SetValue("audioInputSelect", audioInputTableLength)

                elseif NamedControl.GetValue("audioInputSelect") ~= audioInputSelectState then
                    NamedControl.SetText("audioInput",
                        audioInputSources[NamedControl.GetValue("audioInputSelect")].Name)
                    audioInputSelectState = NamedControl.GetValue("audioInputSelect")
                end

                if NamedControl.GetValue("audioOutputSelect") > audioOutputTableLength then
                    NamedControl.SetValue("audioOutputSelect", audioOutputTableLength)

                elseif NamedControl.GetValue("audioOutputSelect") ~= audioOutputSelectState then
                    NamedControl.SetText("audioOutput",
                        audioOutputSources[NamedControl.GetValue("audioOutputSelect")].Name)
                    vaudioOutputSelectState = NamedControl.GetValue("audioOutputSelect")
                end
            end
            if inMeeting ~= inMeetingState then
                if inMeeting == true then
                    for i = 1, 17 do
                        zConfigurationSend(zConfigurationNames[i], zConfiguration[i].Cmd, zConfiguration[i].State)
                    end
                elseif inMeeting == false then
                    for i = 1, 8 do
                        NamedControl.SetText("Participant" .. i, "")
                        NamedControl.SetPosition("muteLED" .. i, 0)
                        NamedControl.SetPosition("videoMuteLED" .. i, 0)
                    end
                end
                inMeetingState = inMeeting

            end
        end
    end


end

Clear()
MyTimer = Timer.New()
MyTimer.EventHandler = TimerClick
MyTimer:Start(.25)
