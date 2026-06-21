import { useEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { Undo, Upload, RefreshCw, CheckCircle, Check } from "lucide-react";

function getAmenityColor(name) {
  const lower = String(name || "").toLowerCase();
  if (lower.includes("vip 1")) return "bg-purple-200 text-purple-800";
  if (lower.includes("vip 2")) return "bg-yellow-200 text-yellow-800";
  if (lower.includes("vip 3")) return "bg-orange-200 text-orange-800";
  if (lower.includes("vip 4")) return "bg-sky-200 text-sky-800";
  if (lower.includes("vip 5")) return "bg-green-200 text-green-800";
  if (lower.includes("vip 6")) return "bg-pink-200 text-pink-800";
  if (lower.includes("vip 7")) return "bg-violet-300 text-violet-900";
  if (lower.includes("kid")) return "bg-pink-200 text-pink-800";
  return "bg-gray-200 text-gray-800";
}

export default function App() {
  const [parsedData, setParsedData] = useState(() => {
    const saved = localStorage.getItem("parsedData");
    return saved ? JSON.parse(saved) : null;
  });

  const [cleanRooms, setCleanRooms] = useState(() => {
    const saved = localStorage.getItem("cleanRooms");
    return saved ? JSON.parse(saved) : [];
  });

  const [deliveredRooms, setDeliveredRooms] = useState(() => {
    const saved = localStorage.getItem("deliveredRooms");
    return saved ? JSON.parse(saved) : [];
  });

  const [deliveryHistory, setDeliveryHistory] = useState(() => {
    const saved = localStorage.getItem("deliveryHistory");
    return saved ? JSON.parse(saved) : [];
  });

  const [roomInput, setRoomInput] = useState("");
  const [viewMode, setViewMode] = useState("ready");
  const [activeFloor, setActiveFloor] = useState("All");
  const [pendingDeliveryRoom, setPendingDeliveryRoom] = useState(null);
  const deliveryTimerRef = useRef(null);

  // Persist states
  useEffect(() => {
    if (parsedData)
      localStorage.setItem("parsedData", JSON.stringify(parsedData));
    else localStorage.removeItem("parsedData");
  }, [parsedData]);

  useEffect(() => {
    localStorage.setItem("cleanRooms", JSON.stringify(cleanRooms));
  }, [cleanRooms]);

  useEffect(() => {
    localStorage.setItem("deliveredRooms", JSON.stringify(deliveredRooms));
  }, [deliveredRooms]);

  useEffect(() => {
    localStorage.setItem("deliveryHistory", JSON.stringify(deliveryHistory));
  }, [deliveryHistory]);

  useEffect(() => {
    return () => {
      if (deliveryTimerRef.current) {
        clearTimeout(deliveryTimerRef.current);
      }
    };
  }, []);

  /**
   * Handles Excel upload and parses Salesforce amenities rows into room-grouped data.
   * Uses defensive parsing to tolerate odd headers, sparse cells, and malformed files.
   * @param {Event} e - File input change event.
   */
  const handleFileUpload = (e) => {
    const file = e?.target?.files?.[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        // Validate XLSX availability before parsing any binary data.
        if (
          typeof XLSX === "undefined" ||
          !XLSX ||
          typeof XLSX.read !== "function"
        ) {
          throw new Error("XLSX library is not loaded.");
        }

        // Read uploaded data as a binary string so SheetJS can handle fake Excel exports.
        const result = event?.target?.result;
        if (!result || typeof result !== "string") {
          throw new Error("No file data was read from the upload.");
        }
        const workbook = XLSX.read(result, { type: "binary" });

        // Resolve the first worksheet from the workbook safely.
        const sheetName = workbook?.SheetNames?.[0];
        if (!sheetName) {
          throw new Error("The uploaded Excel workbook has no sheets.");
        }
        const sheet = workbook?.Sheets?.[sheetName];
        if (!sheet) {
          throw new Error("Unable to access the first sheet in the workbook.");
        }

        // Convert worksheet to a raw 2D array to avoid brittle metadata-driven parsing.
        const rawData = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: "",
        });
        console.log("Raw File Object:", file);
        console.log("Parsed Raw Data:", rawData);
        if (!Array.isArray(rawData) || rawData.length === 0) {
          throw new Error("The uploaded Excel sheet is empty.");
        }

        const isTotalRow = (val) => val && /^\s*total\s*$/i.test(String(val));
        const arrivalTimePattern =
          /Arrival time:\s*(\d{1,2}:\d{2}(?:\s*[AaPp][Mm])?|N\/A)/i;
        const extractArrivalTime = (match) => {
          if (!match || !match[1]) return null;
          return /^n\/a$/i.test(match[1]) ? null : match[1];
        };
        const headerIndex = rawData.findIndex((row) =>
          String(row?.[0] ?? "")
            .toLowerCase()
            .includes("location"),
        );
        if (headerIndex === -1) {
          throw new Error(
            'Could not locate the "Location" column. File format may be incorrect.',
          );
        }

        const parsedRooms = {};

        // Extract row values, parse time/note, and group amenities by location.
        for (let i = headerIndex + 1; i < rawData.length; i++) {
          const row = Array.isArray(rawData[i]) ? rawData[i] : [];

          const location = row[0] ? String(row[0]).trim() : "";
          if (isTotalRow(location)) break;
          if (!location) continue;

          const guestName = row[1] ? String(row[1]).trim() : "";
          const amenityName = row[3] ? String(row[3]).trim() : "";
          const rawDescription = row[4] ? String(row[4]) : "";

          // Pull the arrival time with a safe regex match and preserve null when absent/N/A.
          const timeMatch = rawDescription.match(arrivalTimePattern);
          const time = extractArrivalTime(timeMatch);

          // Remove the arrival-time fragment to leave a clean amenity note.
          const note = rawDescription.replace(arrivalTimePattern, "").trim();

          if (!parsedRooms[location]) {
            parsedRooms[location] = {
              roomNumber: location,
              guestName,
              amenities: [],
            };
          }

          if (amenityName) {
            parsedRooms[location].amenities.push({
              name: amenityName,
              time,
              note,
            });
          }
        }

        // Finalize parsed payload and update state only when we have useful data.
        const finalData = Object.values(parsedRooms);
        if (finalData.length === 0) {
          throw new Error("No valid amenities data found below the header.");
        }

        setParsedData(finalData);
      } catch (error) {
        const message =
          error instanceof Error
            ? error.message
            : "An unexpected error occurred while parsing the Excel file. Please verify the file format and try again.";
        if (error instanceof Error && error.stack) {
          console.error(error.stack);
        } else {
          console.error(error);
        }
        alert(`Parsing Error: ${message}`);
      }
    };

    reader.onerror = (errorEvent) => {
      console.error("File read error:", errorEvent);
      alert(
        "Failed to read the file. Please ensure the file is accessible and try again.",
      );
    };

    reader.readAsBinaryString(file);
  };

  const handleResetDay = () => {
    if (
      window.confirm("Are you sure you want to reset all data for the day?")
    ) {
      if (deliveryTimerRef.current) {
        clearTimeout(deliveryTimerRef.current);
        deliveryTimerRef.current = null;
      }
      setParsedData(null);
      setCleanRooms([]);
      setDeliveredRooms([]);
      setDeliveryHistory([]);
      setViewMode("ready");
      setActiveFloor("All");
      setPendingDeliveryRoom(null);
      localStorage.clear();
    }
  };

  const handleAddCleanRoom = (e) => {
    e.preventDefault();
    const room = roomInput.trim();
    if (room && !cleanRooms.includes(room)) {
      setCleanRooms([...cleanRooms, room]);
    }
    setRoomInput("");
  };

  const handleMarkDelivered = (roomNumber) => {
    if (pendingDeliveryRoom) return;
    if (deliveredRooms.includes(roomNumber)) return;

    const confirmed = window.confirm(
      `Confirm delivery for room ${roomNumber}?`,
    );
    if (!confirmed) return;

    setPendingDeliveryRoom(roomNumber);

    if (deliveryTimerRef.current) {
      clearTimeout(deliveryTimerRef.current);
    }
    deliveryTimerRef.current = setTimeout(() => {
      setDeliveredRooms((current) =>
        current.includes(roomNumber) ? current : [...current, roomNumber],
      );
      setDeliveryHistory((current) => [...current, roomNumber]);
      setPendingDeliveryRoom(null);
      deliveryTimerRef.current = null;
    }, 500);
  };

  const handleUndoDelivery = () => {
    if (deliveryHistory.length === 0) return;
    const historyCopy = [...deliveryHistory];
    const lastDelivered = historyCopy.pop();

    setDeliveryHistory(historyCopy);
    setDeliveredRooms(deliveredRooms.filter((r) => r !== lastDelivered));
  };

  const floors = ["All", "M", "1", "2", "3", "4", "5"];

  const getRoomFloor = (roomNumber) => {
    const normalizedRoom = String(roomNumber || "").trim();
    if (!normalizedRoom) return null;
    if (normalizedRoom.charAt(0) === "0") return "M";

    const leadingDigit = normalizedRoom.charAt(0);
    if (["1", "2", "3", "4", "5"].includes(leadingDigit)) {
      return leadingDigit;
    }

    return null;
  };

  const roomsByView = parsedData
    ? viewMode === "all"
      ? parsedData
      : parsedData.filter(
          (room) =>
            cleanRooms.includes(room.roomNumber) &&
            !deliveredRooms.includes(room.roomNumber),
        )
    : [];

  const visibleRooms = roomsByView.filter((room) => {
    const roomFloor = getRoomFloor(room?.roomNumber);
    if (!roomFloor) return false;
    if (activeFloor === "All") return true;
    return roomFloor === activeFloor;
  });

  return (
    <div className="min-h-screen bg-gray-50 pb-20">
      {/* Top Header */}
      <div className="bg-white shadow-sm p-4 flex justify-between items-center sticky top-0 z-20">
        <h1 className="text-xl font-bold text-gray-800">Hotel Amenities</h1>
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={() =>
              setViewMode((current) => (current === "ready" ? "all" : "ready"))
            }
            className="text-gray-700 flex items-center text-sm font-medium border border-gray-200 px-3 py-1.5 rounded-md hover:bg-gray-50"
          >
            {viewMode === "ready"
              ? "View: Ready to Deliver"
              : "View: All Amenities"}
          </button>
          <button
            onClick={handleResetDay}
            className="text-red-600 flex items-center text-sm font-medium border border-red-200 px-3 py-1.5 rounded-md hover:bg-red-50"
          >
            <RefreshCw className="w-4 h-4 mr-2" />
            Reset Day
          </button>
        </div>
      </div>

      <div className="max-w-md mx-auto p-4">
        {/* Upload Section */}
        {!parsedData && (
          <div className="bg-white p-8 rounded-xl shadow-sm border border-gray-200 text-center mb-6">
            <Upload className="w-12 h-12 text-blue-500 mx-auto mb-4" />
            <h2 className="text-lg font-semibold mb-2">
              Upload Amenities File
            </h2>
            <p className="text-gray-500 text-sm mb-4">
              Select the Salesforce export (.xlsx)
            </p>
            <label className="bg-blue-600 text-white px-6 py-2 rounded-lg font-medium cursor-pointer hover:bg-blue-700 transition-colors inline-block">
              Choose File
              <input
                type="file"
                accept=".xlsx, .xls"
                onChange={handleFileUpload}
                className="hidden"
              />
            </label>
          </div>
        )}

        {/* Controls Section (Sticky when parsedData exists) */}
        {parsedData && (
          <div className="bg-white p-4 rounded-xl shadow-md border border-gray-200 mb-6 sticky top-18 z-10">
            <form onSubmit={handleAddCleanRoom} className="flex gap-2 mb-4">
              <input
                type="text"
                inputMode="numeric"
                pattern="[0-9]*"
                value={roomInput}
                onChange={(e) => setRoomInput(e.target.value)}
                placeholder="Enter room number..."
                className="flex-1 border border-gray-300 rounded-lg px-4 py-3 text-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
              />
              <button
                type="submit"
                className="bg-blue-600 text-white px-6 py-3 rounded-lg font-medium hover:bg-blue-700 transition-colors whitespace-nowrap"
              >
                Add Clean
              </button>
            </form>

            <div className="flex justify-between items-center">
              <div className="text-sm font-medium text-gray-600">
                Delivered:{" "}
                <span className="text-green-600 font-bold">
                  {deliveredRooms.length}
                </span>{" "}
                / Total Clean:{" "}
                <span className="text-blue-600 font-bold">
                  {cleanRooms.length}
                </span>
              </div>

              <button
                onClick={handleUndoDelivery}
                disabled={deliveryHistory.length === 0}
                className="flex items-center text-sm font-medium text-gray-700 disabled:text-gray-300 disabled:cursor-not-allowed bg-gray-100 px-3 py-2 rounded-lg hover:bg-gray-200 transition-colors"
              >
                <Undo className="w-4 h-4 mr-1" />
                Undo Last
              </button>
            </div>
          </div>
        )}

        {/* Main Body (Compact List) */}
        {parsedData && (
          <div>
            <div className="flex flex-wrap gap-2 mb-2">
              {floors.map((floor) => (
                <button
                  key={floor}
                  type="button"
                  onClick={() => setActiveFloor(floor)}
                  className={`px-3 py-1.5 rounded-full text-xs font-semibold border transition-colors ${activeFloor === floor ? "bg-blue-600 text-white border-blue-600" : "bg-white text-gray-700 border-gray-200 hover:bg-gray-50"}`}
                >
                  {floor === "All"
                    ? "All"
                    : floor === "M"
                      ? "M"
                      : `Fl ${floor}`}
                </button>
              ))}
            </div>

            {visibleRooms.length === 0 && (
              <div className="text-center p-8 bg-gray-50 border-2 border-dashed border-gray-200 rounded-xl">
                <CheckCircle className="w-12 h-12 text-gray-300 mx-auto mb-3" />
                <p className="text-gray-500 font-medium">
                  {viewMode === "all"
                    ? "No rooms match the current floor filter."
                    : "No clean rooms pending delivery."}
                </p>
                <p className="text-gray-400 text-sm mt-1">
                  {viewMode === "all"
                    ? "Switch floors or return to the ready view."
                    : "Add a clean room above to see amenities."}
                </p>
              </div>
            )}

            <div className="bg-white rounded-lg border border-gray-200 overflow-hidden divide-y divide-gray-100">
              {visibleRooms.map((room) => {
                const isDelivered = deliveredRooms.includes(room.roomNumber);
                const isPending = pendingDeliveryRoom === room.roomNumber;
                const primaryAmenity = room.amenities[0];
                const displayTime =
                  room.amenities.find((a) => a.time)?.time ?? null;
                const displayNote =
                  room.amenities
                    .map((a) => a.note)
                    .filter(Boolean)
                    .join(" · ") || null;

                return (
                  <div
                    key={room.roomNumber}
                    className={`flex items-center gap-2 p-2 border-b border-gray-100 last:border-b-0 transition-all ${
                      isDelivered
                        ? "bg-green-50 opacity-60"
                        : isPending
                          ? "opacity-50"
                          : "bg-white"
                    }`}
                  >
                    <div className="shrink-0 w-18">
                      <div
                        className={`font-bold text-lg leading-tight ${isDelivered ? "line-through text-gray-500" : "text-gray-900"}`}
                      >
                        {room.roomNumber}
                      </div>
                      <div className="flex flex-wrap gap-0.5 mt-0.5">
                        {room.amenities.map((amenity, idx) => (
                          <span
                            key={idx}
                            className={`px-1 py-0.5 rounded text-[10px] font-bold uppercase leading-none ${getAmenityColor(amenity.name)}`}
                          >
                            {amenity.name || "?"}
                          </span>
                        ))}
                      </div>
                    </div>

                    <div className="flex-1 min-w-0">
                      {displayTime && (
                        <div className="text-xs font-semibold text-red-600 leading-tight">
                          {displayTime}
                        </div>
                      )}
                      {displayNote && (
                        <div className="text-xs text-gray-500 leading-tight truncate">
                          {displayNote}
                        </div>
                      )}
                      {!displayTime && !displayNote && primaryAmenity && (
                        <div className="text-xs text-gray-400 leading-tight">
                          {room.guestName || "—"}
                        </div>
                      )}
                    </div>

                    <button
                      type="button"
                      onClick={() => handleMarkDelivered(room.roomNumber)}
                      disabled={isDelivered || isPending}
                      aria-label={
                        isDelivered
                          ? `Room ${room.roomNumber} delivered`
                          : `Mark room ${room.roomNumber} delivered`
                      }
                      className={`shrink-0 w-9 h-9 rounded-full flex items-center justify-center transition-colors ${
                        isDelivered
                          ? "bg-green-200 text-green-700 cursor-default"
                          : isPending
                            ? "bg-gray-200 text-gray-400 cursor-not-allowed"
                            : "bg-green-500 hover:bg-green-600 text-white"
                      }`}
                    >
                      <Check className="w-5 h-5" />
                    </button>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
