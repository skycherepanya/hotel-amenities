import { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Undo, Upload, RefreshCw, CheckCircle } from 'lucide-react';

function getAmenityColor(name) {
  const lower = (name || '').toLowerCase();
  if (lower.includes('vip 1')) return 'bg-purple-200 text-purple-800';
  if (lower.includes('vip 2')) return 'bg-yellow-200 text-yellow-800';
  if (lower.includes('vip 3')) return 'bg-orange-200 text-orange-800';
  if (lower.includes('vip 4')) return 'bg-sky-200 text-sky-800';
  if (lower.includes('vip 5')) return 'bg-green-200 text-green-800';
  if (lower.includes('vip 6')) return 'bg-pink-200 text-pink-800';
  if (lower.includes('vip 7')) return 'bg-violet-300 text-violet-900';
  if (lower.includes("kid's") || lower.includes("kids")) return 'bg-pink-200 text-pink-800';
  return 'bg-gray-200 text-gray-800';
}

export default function App() {
  const [parsedData, setParsedData] = useState(() => {
    const saved = localStorage.getItem('parsedData');
    return saved ? JSON.parse(saved) : null;
  });
  
  const [cleanRooms, setCleanRooms] = useState(() => {
    const saved = localStorage.getItem('cleanRooms');
    return saved ? JSON.parse(saved) : [];
  });
  
  const [deliveredRooms, setDeliveredRooms] = useState(() => {
    const saved = localStorage.getItem('deliveredRooms');
    return saved ? JSON.parse(saved) : [];
  });
  
  const [deliveryHistory, setDeliveryHistory] = useState(() => {
    const saved = localStorage.getItem('deliveryHistory');
    return saved ? JSON.parse(saved) : [];
  });

  const [roomInput, setRoomInput] = useState('');

  // Persist states
  useEffect(() => {
    if (parsedData) localStorage.setItem('parsedData', JSON.stringify(parsedData));
    else localStorage.removeItem('parsedData');
  }, [parsedData]);

  useEffect(() => {
    localStorage.setItem('cleanRooms', JSON.stringify(cleanRooms));
  }, [cleanRooms]);

  useEffect(() => {
    localStorage.setItem('deliveredRooms', JSON.stringify(deliveredRooms));
  }, [deliveredRooms]);

  useEffect(() => {
    localStorage.setItem('deliveryHistory', JSON.stringify(deliveryHistory));
  }, [deliveryHistory]);

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
        if (typeof XLSX === 'undefined' || !XLSX || typeof XLSX.read !== 'function') {
          throw new Error('XLSX library is not loaded.');
        }

        // Read uploaded bytes from FileReader and ensure we received a valid buffer.
        const result = event?.target?.result;
        if (!result) {
          throw new Error('No file data was read from the upload.');
        }
        const data = new Uint8Array(result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Resolve the first worksheet from the workbook safely.
        const sheetName = workbook?.SheetNames?.[0];
        if (!sheetName) {
          throw new Error('The uploaded Excel workbook has no sheets.');
        }
        const sheet = workbook?.Sheets?.[sheetName];
        if (!sheet) {
          throw new Error('Unable to access the first sheet in the workbook.');
        }

        // Convert worksheet to a raw 2D array to avoid brittle metadata-driven parsing.
        const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        if (!Array.isArray(rawData) || rawData.length === 0) {
          throw new Error('The uploaded Excel sheet is empty.');
        }

        // Fuzzy "Location" header detection to handle casing/spacing/hidden-char variations.
        const isHeader = (val) => val && /^\s*location\s*$/i.test(String(val));
        const headerIndex = rawData.findIndex((row) => isHeader(row?.[0]));
        if (headerIndex === -1) {
          throw new Error('Could not locate the "Location" column. File format may be incorrect.');
        }

        const parsedRooms = {};

        // Extract row values, parse time/note, and group amenities by location.
        for (let i = headerIndex + 1; i < rawData.length; i++) {
          const row = Array.isArray(rawData[i]) ? rawData[i] : [];

          const location = row[0] ? String(row[0]).trim() : '';
          if (/^\s*total\s*$/i.test(location)) break;
          if (!location) continue;

          const guestName = row[1] ? String(row[1]).trim() : '';
          const amenityName = row[3] ? String(row[3]).trim() : '';
          const rawDescription = row[4] ? String(row[4]) : '';

          // Pull the arrival time with a safe regex match and preserve null when absent/N/A.
          const timeMatch = rawDescription.match(/Arrival time:\s*(\d{1,2}:\d{2}(?:\s*[AaPp][Mm])?|N\/A)/i);
          const time = timeMatch && timeMatch[1] && timeMatch[1].toUpperCase() !== 'N/A' ? timeMatch[1] : null;

          // Remove the arrival-time fragment to leave a clean amenity note.
          const note = rawDescription
            .replace(/Arrival time:\s*(\d{1,2}:\d{2}(?:\s*[AaPp][Mm])?|N\/A)\s*/i, '')
            .trim();

          if (!parsedRooms[location]) {
            parsedRooms[location] = {
              roomNumber: location,
              guestName,
              amenities: []
            };
          }

          if (amenityName) {
            parsedRooms[location].amenities.push({
              name: amenityName,
              time,
              note
            });
          }
        }

        // Finalize parsed payload and update state only when we have useful data.
        const finalData = Object.values(parsedRooms);
        if (finalData.length === 0) {
          throw new Error('No valid amenities data found below the header.');
        }

        setParsedData(finalData);
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Unknown parsing error';
        console.error('Excel Parsing Error:', error);
        if (error instanceof Error && error.stack) {
          console.error(error.stack);
        }
        alert(`Parsing Error: ${message}`);
      }
    };

    reader.onerror = () => {
      alert('Failed to read the file. It might be corrupted.');
    };

    reader.readAsArrayBuffer(file);
  };

  const handleResetDay = () => {
    if (window.confirm('Are you sure you want to reset all data for the day?')) {
      setParsedData(null);
      setCleanRooms([]);
      setDeliveredRooms([]);
      setDeliveryHistory([]);
      localStorage.clear();
    }
  };

  const handleAddCleanRoom = (e) => {
    e.preventDefault();
    const room = roomInput.trim();
    if (room && !cleanRooms.includes(room)) {
      setCleanRooms([...cleanRooms, room]);
    }
    setRoomInput('');
  };

  const handleMarkDelivered = (roomNumber) => {
    if (!deliveredRooms.includes(roomNumber)) {
      setDeliveredRooms([...deliveredRooms, roomNumber]);
      setDeliveryHistory([...deliveryHistory, roomNumber]);
    }
  };

  const handleUndoDelivery = () => {
    if (deliveryHistory.length === 0) return;
    const historyCopy = [...deliveryHistory];
    const lastDelivered = historyCopy.pop();
    
    setDeliveryHistory(historyCopy);
    setDeliveredRooms(deliveredRooms.filter(r => r !== lastDelivered));
  };

  // Filter visible cards
  const visibleRooms = parsedData 
    ? parsedData.filter(room => cleanRooms.includes(room.roomNumber) && !deliveredRooms.includes(room.roomNumber))
    : [];

  return (
    <div className="min-h-screen bg-gray-50 pb-20">
      {/* Top Header */}
      <div className="bg-white shadow-sm p-4 flex justify-between items-center sticky top-0 z-20">
        <h1 className="text-xl font-bold text-gray-800">Hotel Amenities</h1>
        <button 
          onClick={handleResetDay}
          className="text-red-600 flex items-center text-sm font-medium border border-red-200 px-3 py-1.5 rounded-md hover:bg-red-50"
        >
          <RefreshCw className="w-4 h-4 mr-2" />
          Reset Day
        </button>
      </div>

      <div className="max-w-md mx-auto p-4">
        {/* Upload Section */}
        {!parsedData && (
          <div className="bg-white p-8 rounded-xl shadow-sm border border-gray-200 text-center mb-6">
            <Upload className="w-12 h-12 text-blue-500 mx-auto mb-4" />
            <h2 className="text-lg font-semibold mb-2">Upload Amenities File</h2>
            <p className="text-gray-500 text-sm mb-4">Select the Salesforce export (.xlsx)</p>
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
          <div className="bg-white p-4 rounded-xl shadow-md border border-gray-200 mb-6 sticky top-[72px] z-10">
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
                Delivered: <span className="text-green-600 font-bold">{deliveredRooms.length}</span> / 
                Total Clean: <span className="text-blue-600 font-bold">{cleanRooms.length}</span>
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

        {/* Main Body (Card List) */}
        {parsedData && (
          <div className="space-y-4">
            {visibleRooms.length === 0 && (
              <div className="text-center p-8 bg-gray-50 border-2 border-dashed border-gray-200 rounded-xl">
                <CheckCircle className="w-12 h-12 text-gray-300 mx-auto mb-3" />
                <p className="text-gray-500 font-medium">No clean rooms pending delivery.</p>
                <p className="text-gray-400 text-sm mt-1">Add a clean room above to see amenities.</p>
              </div>
            )}

            {visibleRooms.map((room) => (
              <div key={room.roomNumber} className="bg-white rounded-xl shadow-md border border-gray-200 overflow-hidden">
                <div className="p-4 border-b border-gray-100">
                  <div className="flex justify-between items-baseline">
                    <h2 className="text-3xl font-bold text-gray-900">{room.roomNumber}</h2>
                    <span className="text-lg font-medium text-gray-600">{room.guestName}</span>
                  </div>
                </div>
                
                <div className="p-4 space-y-4">
                  {room.amenities.map((amenity, idx) => (
                    <div key={idx} className="flex flex-col gap-2 pb-4 border-b border-gray-50 last:border-0 last:pb-0">
                      <div className="flex justify-between items-start">
                        <span className={`px-3 py-1 rounded-full text-sm font-bold uppercase tracking-wide ${getAmenityColor(amenity.name)}`}>
                          {amenity.name || 'Unknown Amenity'}
                        </span>
                        {amenity.time && (
                          <span className="text-red-600 font-bold text-lg">{amenity.time}</span>
                        )}
                      </div>
                      
                      {amenity.note && (
                        <div className="bg-gray-50 p-3 rounded-lg text-sm text-gray-600 italic">
                          {amenity.note}
                        </div>
                      )}
                    </div>
                  ))}
                </div>
                
                <button 
                  onClick={() => handleMarkDelivered(room.roomNumber)}
                  className="w-full bg-green-500 hover:bg-green-600 text-white font-bold text-lg py-4 transition-colors flex items-center justify-center"
                >
                  <CheckCircle className="w-6 h-6 mr-2" />
                  Mark Delivered
                </button>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
