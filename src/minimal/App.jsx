import React, { useEffect, useState } from 'react';
import { 
  getSelectedText, 
  insertAtCursor, 
  replaceSelectionWithHtml, 
  replaceSelectionAsTracked, 
  getTrackChangesState, 
  setTrackChangesState,
  replaceSelectionWithNativeTrackedRevisions,
  acceptAllTrackedChanges,
  acceptTrackedChangesInRange
} from './wordUtils';
import { trackChangesDiffPreview, computeWordLevelDiff, generateRevisionOoxml } from './diff';
import TestRunner from './TestRunner';

// Complex sample modeled after typical patent drafting paragraphs with dense edits
const OLD_TEXT = '[0075] FIG. 1 illustrates an illustration of a platform consistent with various embodiments. By way of non-limiting example, the online platform (100-for-facilitating) may facilitate electronic signing and witnessing of electronic signing. The platform (100) may be hosted on a centralized server (110), such as, for example, which may be implemented using a cloud computing service. The centralized server (110) may communicate with other various network entities (such as a smartphone laptop, a tablet computer etc. computers) and other electronic devices (such as, desktop computers, etc) over, via a communication network (112), such as, but not limited to, the Internet.';

const NEW_TEXT = '[0075] FIG. 1 illustrates a platform consistent with various embodiments. By way of example, the online platform (100) facilitates electronic signing and witnessing of documents. The platform (100) may be hosted on a centralized server (110), for example implemented using a cloud computing service. The centralized server (110) communicates with network entities (e.g., smartphones, laptops, tablet computers, and desktop computers) via a communication network (112), such as the Internet.';

export default function App() {
  const [status, setStatus] = useState('Ready');
  const [oldText, setOldText] = useState(OLD_TEXT);
  const [newText, setNewText] = useState(NEW_TEXT);
  const [trackOn, setTrackOn] = useState(null); // null = unknown/unsupported
  const [usingOoxml, setUsingOoxml] = useState(true); // Whether to use OOXML revisions
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingProgress, setProcessingProgress] = useState({ current: 0, total: 0 });
  const [showTests, setShowTests] = useState(false); // Toggle test runner view

  useEffect(() => {
    (async () => {
      const state = await getTrackChangesState();
      setTrackOn(state);
    })();
  }, []);

  const handleInsertOld = async () => {
    setStatus('Inserting…');
    try {
      await insertAtCursor(oldText);
      setStatus('Old text inserted.');
    } catch (e) {
      console.error(e);
      setStatus('Failed to insert text.');
    }
  };

  const handlePreview = async () => {
    setStatus('Generating preview with tracked changes…');
    try {
      const selected = await getSelectedText();
      
      if (!selected || selected.trim().length === 0) {
        setStatus('Please select the old text first.');
        return;
      }
      
      if (usingOoxml) {
        // Save current tracking state
        const wasTrackingOn = trackOn;
        
        // Generate diffs
        const diffs = computeWordLevelDiff(selected, newText);
        console.log('Generated diffs:', diffs);
        
        // Set up progress tracking
        setIsProcessing(true);
        setProcessingProgress({ current: 0, total: diffs.length });
        
        // Progress callback for incremental updates
        const onProgress = (current, total, message) => {
          setProcessingProgress({ current, total });
          setStatus(`Processing: ${message || `${current}/${total} diffs`}`);
        };
        
        try {
          // Use OOXML to insert tracked changes incrementally
          await replaceSelectionWithNativeTrackedRevisions(diffs, wasTrackingOn, onProgress);
          
          // Restore tracking state in UI
          if (wasTrackingOn !== null) {
            await setTrackChangesState(wasTrackingOn);
            const newState = await getTrackChangesState();
            setTrackOn(newState);
          }
          
          // Success message
          setStatus('✅ Preview applied using OOXML. Tracked changes inserted incrementally.');
        } catch (error) {
          setStatus(`❌ Error: ${error.message}`);
          throw error;
        } finally {
          setIsProcessing(false);
          setProcessingProgress({ current: 0, total: 0 });
        }
      } else {
        // Fallback: visual HTML preview
        const html = trackChangesDiffPreview(selected, newText);
        await replaceSelectionWithHtml(html);
        setStatus('Preview applied (visual only).');
      }
    } catch (e) {
      console.error('Preview error:', e);
      console.error('Error details:', {
        message: e.message,
        name: e.name,
        stack: e.stack,
        errorCode: e.errorCode,
        debugInfo: e.debugInfo
      });
      
      setStatus(`❌ Error: ${e.message || e.toString()}`);
    }
  };

  const handleApplyTracked = async () => {
    setStatus('Applying tracked replace…');
    try {
      await replaceSelectionAsTracked(newText);
      setStatus('Applied as single tracked change.');
    } catch (e) {
      console.error(e);
      setStatus('Failed to apply tracked change.');
    }
  };

  const toggleTrack = async () => {
    const desired = !(!!trackOn);
    setStatus(desired ? 'Turning Track Changes ON…' : 'Turning Track Changes OFF…');
    const result = await setTrackChangesState(desired);
    if (result === null) {
      setStatus('Track Changes toggle not supported in this environment.');
    } else {
      setTrackOn(result);
      setStatus(`Track Changes ${result ? 'ON' : 'OFF'}.`);
    }
  };

  const handleAcceptAll = async () => {
    setStatus('Accepting all tracked changes…');
    try {
      const result = await acceptAllTrackedChanges();
      if (result.success) {
        setStatus(`✅ Accepted ${result.accepted === -1 ? 'all' : result.accepted} tracked changes`);
      } else {
        setStatus(`❌ Error: ${result.error}`);
      }
    } catch (e) {
      console.error('Accept all error:', e);
      setStatus(`❌ Error: ${e.message}`);
    }
  };

  const handleAcceptSelection = async () => {
    setStatus('Accepting tracked changes in selection…');
    try {
      const result = await acceptTrackedChangesInRange('selection');
      if (result.success) {
        setStatus(`✅ Accepted ${result.accepted} tracked changes in selection`);
      } else {
        setStatus(`❌ Error: ${result.error}`);
      }
    } catch (e) {
      console.error('Accept selection error:', e);
      setStatus(`❌ Error: ${e.message}`);
    }
  };

  return (
    <div style={{ padding: 12 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div>
          <h2 style={{ margin: 0 }}>Example — Track Changes Sample</h2>
          <p style={{ color: '#555', marginTop: 4 }}>Insert complex text, preview visual diffs, or apply a single tracked change.</p>
        </div>
        <button 
          onClick={() => setShowTests(!showTests)}
          style={{ 
            padding: '6px 12px',
            background: showTests ? '#0078d4' : '#f0f0f0',
            color: showTests ? 'white' : '#333',
            border: '1px solid #ccc',
            borderRadius: 4,
            cursor: 'pointer'
          }}
        >
          {showTests ? 'Hide Tests' : 'Show Tests'}
        </button>
      </div>

      {showTests && (
        <div style={{ 
          marginTop: 12, 
          marginBottom: 12,
          padding: 12,
          background: '#f9f9f9',
          borderRadius: 6,
          border: '1px solid #ddd'
        }}>
          <TestRunner />
        </div>
      )}

      <div style={{ display: 'grid', gap: 12 }}>
        <div>
          <label style={{ fontWeight: 600 }}>Old text (to insert)</label>
          <textarea value={oldText} onChange={(e) => setOldText(e.target.value)} rows={6} style={{ width: '100%', boxSizing: 'border-box' }} />
          <button onClick={handleInsertOld} style={{ marginTop: 8 }}>Insert Old Text</button>
        </div>

        <div>
          <label style={{ fontWeight: 600 }}>New text (to preview/apply)</label>
          <textarea value={newText} onChange={(e) => setNewText(e.target.value)} rows={6} style={{ width: '100%', boxSizing: 'border-box' }} />
          <div style={{ display: 'flex', gap: 8, marginTop: 8, alignItems: 'center', flexWrap: 'wrap' }}>
            <button onClick={handlePreview} style={{ fontWeight: 600 }}>
              Preview Changes
            </button>
            <button onClick={handleApplyTracked} disabled={!trackOn}>Apply As Tracked Change</button>
            <span style={{ padding: '4px 8px', borderRadius: 6, background: trackOn ? '#e3fcef' : '#ffebe6', color: trackOn ? '#006644' : '#bf2600' }}>
              Track Changes: {trackOn === null ? 'Unknown/Unsupported' : trackOn ? 'On' : 'Off'}
            </span>
            <button onClick={toggleTrack}>{trackOn ? 'Turn Off' : 'Turn On'} Track Changes</button>
          </div>
          <div style={{ marginTop: 4, fontSize: 12, color: '#666' }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <input 
                type="checkbox" 
                checked={usingOoxml} 
                onChange={(e) => setUsingOoxml(e.target.checked)}
              />
              Use native Word tracking (real interactive tracked changes with hover Accept/Reject)
            </label>
          </div>
        </div>

        <div style={{ marginTop: 12, padding: 12, background: '#f9f9f9', borderRadius: 6 }}>
          <div style={{ fontWeight: 600, marginBottom: 8 }}>Tracked Change Management</div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <button 
              onClick={handleAcceptSelection}
              style={{ padding: '6px 12px', background: '#0078d4', color: 'white', border: 'none', borderRadius: 4, cursor: 'pointer' }}
            >
              Accept in Selection
            </button>
            <button 
              onClick={handleAcceptAll}
              style={{ padding: '6px 12px', background: '#d13438', color: 'white', border: 'none', borderRadius: 4, cursor: 'pointer' }}
            >
              Accept All in Document
            </button>
          </div>
          <div style={{ fontSize: 11, color: '#666', marginTop: 4 }}>
            Use these to accept tracked changes created by "Preview Changes" operation
          </div>
        </div>

        <div style={{ color: '#666', marginTop: 12 }}>
          Status: {status}
          {isProcessing && (
            <div style={{ marginTop: 8, padding: 12, background: '#f0f0f0', borderRadius: 6 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                <div style={{ 
                  width: 20, 
                  height: 20, 
                  border: '3px solid #e3e3e3',
                  borderTop: '3px solid #0078d4',
                  borderRadius: '50%',
                  animation: 'spin 1s linear infinite'
                }} />
                <span style={{ fontWeight: 600 }}>Processing tracked changes...</span>
              </div>
              {processingProgress.total > 0 && (
                <div style={{ marginTop: 4 }}>
                  <div style={{ 
                    width: '100%', 
                    height: 6, 
                    background: '#e3e3e3', 
                    borderRadius: 3,
                    overflow: 'hidden'
                  }}>
                    <div style={{
                      width: `${(processingProgress.current / processingProgress.total) * 100}%`,
                      height: '100%',
                      background: '#0078d4',
                      transition: 'width 0.3s ease'
                    }} />
                  </div>
                  <div style={{ fontSize: 11, color: '#666', marginTop: 4 }}>
                    {processingProgress.current} / {processingProgress.total} diffs processed
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
        <div style={{ fontSize: 12, color: '#6b778c' }}>
          Tip: Insert the old text, select it, then Preview for tracked changes preview. 
          {usingOoxml && ' With OOXML enabled, hovering over changes shows Accept/Reject options.'}
        </div>
        <style>{`
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        `}</style>
      </div>
    </div>
  );
}