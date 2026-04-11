'use client';

import { useState } from 'react';

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');
  const [progress, setProgress] = useState(0);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile && selectedFile.name.endsWith('.md')) {
      setFile(selectedFile);
      setError('');
    } else {
      setError('Please select a valid .md file');
      setFile(null);
    }
  };

  const handleUpload = async () => {
    if (!file) return;

    setLoading(true);
    setMessage('Uploading file...');
    setError('');
    setProgress(0);

    try {
      const formData = new FormData();
      formData.append('file', file);

      // 1. Upload and get the Job ID Ticket
      const uploadRes = await fetch('http://localhost:7860/upload', {
        method: 'POST',
        body: formData,
      });

      if (!uploadRes.ok) throw new Error('Upload failed');
      const uploadData = await uploadRes.json();
      
      if (uploadData.status === 'error') throw new Error(uploadData.message);
      
      const jobId = uploadData.job_id;

      // 2. Start polling the status endpoint
      let isComplete = false;
      while (!isComplete) {
        // Wait 2 seconds between checks
        await new Promise((resolve) => setTimeout(resolve, 2000));
        
        const statusRes = await fetch(`http://localhost:7860/status/${jobId}`);
        const statusData = await statusRes.json();

        if (statusData.status === 'error') {
          throw new Error(statusData.message);
        }

        // Update the UI with what the backend is currently doing
        setMessage(statusData.message);
        setProgress(statusData.progress);

        if (statusData.status === 'completed') {
          isComplete = true;
          console.log("FINAL DATA:", statusData.data); // Your JSON is here!
          setMessage('Success! Presentation generated (check console for data).');
          setFile(null);
          setLoading(false);
        }
      }

    } catch (err) {
      setError(err instanceof Error ? err.message : 'Pipeline failed');
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-gray-50 py-12 px-4 sm:px-6 lg:px-8">
      <div className="max-w-md w-full bg-white rounded-lg shadow-md p-8">
        <h1 className="text-3xl font-bold text-center mb-2 text-gray-900">Markdown to PPTX</h1>
        <p className="text-center text-gray-600 mb-8">AI-Powered Presentation Generator</p>

        <div className="space-y-4">
          <div className={`border-2 border-dashed rounded-lg p-6 transition ${loading ? 'border-gray-200 opacity-50' : 'border-gray-300 hover:border-blue-500'}`}>
            <label className={`block ${loading ? 'cursor-not-allowed' : 'cursor-pointer'}`}>
              <input type="file" accept=".md" onChange={handleFileChange} className="hidden" disabled={loading} />
              <div className="text-center">
                <p className="mt-2 text-sm text-gray-600">
                  {file ? <span className="text-blue-600 font-semibold">{file.name}</span> : 'Click to upload .md file'}
                </p>
              </div>
            </label>
          </div>

          {error && <div className="bg-red-50 border border-red-200 rounded p-3 text-red-700 text-sm">{error}</div>}

          {/* DYNAMIC PROGRESS BAR UI */}
          {loading && (
            <div className="space-y-2 mt-4">
              <div className="flex justify-between text-sm font-medium text-gray-700">
                <span>{message}</span>
                <span>{progress}%</span>
              </div>
              <div className="w-full bg-gray-200 rounded-full h-2.5">
                <div 
                  className="bg-blue-600 h-2.5 rounded-full transition-all duration-500 ease-out" 
                  style={{ width: `${progress}%` }}
                ></div>
              </div>
            </div>
          )}

          {!loading && message && !error && (
            <div className="bg-green-50 border border-green-200 rounded p-3 text-green-700 text-sm font-medium">
              {message}
            </div>
          )}

          <button
            onClick={handleUpload}
            disabled={!file || loading}
            className="w-full bg-blue-600 text-white py-2 rounded-lg font-semibold hover:bg-blue-700 disabled:opacity-50 transition mt-4"
          >
            {loading ? 'Processing...' : 'Generate AI Presentation'}
          </button>
        </div>
      </div>
    </div>
  );
}