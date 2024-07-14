import 'bootstrap/dist/css/bootstrap.min.css';
import React, { useState, useEffect } from 'react';
import axios from 'axios';
import './App.css';
import { Button, Container, Form, Col, Row, Modal } from 'react-bootstrap';

function App() {
  const [ms6File, setMs6File] = useState(null);
  const [bmsFile, setBmsFile] = useState(null);
  const [status, setStatus] = useState('');
  const [loading, setLoading] = useState(false);
  const [showModal, setShowModal] = useState(false);

  const handleFileChange = (e) => {
    if (e.target.name === 'ms6File') {
      setMs6File(e.target.files[0]);
    } else if (e.target.name === 'bmsFile') {
      setBmsFile(e.target.files[0]);
    }
  };

  const pollStatus = async () => {
    try {
      const response = await axios.get('http://127.0.0.1:5000/status');
      setStatus(response.data.message);
    } catch (error) {
      console.error("Error fetching status", error);
    }
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setLoading(true);
    const formData = new FormData();
    formData.append('ms6File', ms6File);
    formData.append('bmsFile', bmsFile);

    try {
      const response = await axios.post('http://127.0.0.1:5000/generate-certificates', formData, {
        responseType: 'blob',
      });
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'certificates.pdf');
      document.body.appendChild(link);
      link.click();
      link.remove();

      // Delete files after downloading
      await axios.post('http://127.0.0.1:5000/delete-files');
      console.log("Files deleted successfully");

      setTimeout(() => {
        setShowModal(true); // Show modal after download and a delay
      }, 3000);
    } catch (error) {
      console.error('Error generating certificates', error);
    } finally {
      setLoading(false);
    }
  };

  const handleCloseModal = () => {
    setShowModal(false);
  };

  useEffect(() => {
    if (!loading && showModal) {
      console.log('Modal should be shown'); // Debugging statement
    }
  }, [loading, showModal]);

  useEffect(() => {
    const intervalId = setInterval(pollStatus, 1000); // Poll status every second
    return () => clearInterval(intervalId);
  }, []);

  return (
    <Container className="app-container">
      <Row className="justify-content-md-center">
        <Col md="auto">
          {!loading && <h1 className="app-title">Certificate Generator</h1>}
          {loading ? (
            <div className="spinner-container">
              <svg xmlns="http://www.w3.org/2000/svg" width="300" height="300">
                <circle id="arc1" className="circle" cx="150" cy="150" r="120" opacity=".89" fill="none" stroke="#632b26" strokeWidth="12" strokeLinecap="square" strokeOpacity=".99213" paintOrder="fill markers stroke" />
                <circle id="arc2" className="circle" cx="150" cy="150" r="120" opacity=".49" fill="none" stroke="#632b26" strokeWidth="8" strokeLinecap="square" strokeOpacity=".99213" paintOrder="fill markers stroke" />
                <circle id="arc3" className="circle" cx="150" cy="150" r="100" opacity=".49" fill="none" stroke="#632b26" strokeWidth="20" strokeLinecap="square" strokeOpacity=".99213" paintOrder="fill markers stroke" />
                <circle id="arc4" className="circle" cx="150" cy="150" r="120" opacity=".49" fill="none" stroke="#632b26" strokeWidth="30" strokeLinecap="square" strokeOpacity=".99213" paintOrder="fill markers stroke" />
                <circle id="arc5" className="circle" cx="150" cy="150" r="100" opacity=".89" fill="none" stroke="#632b26" strokeWidth="8" strokeLinecap="square" strokeOpacity=".99213" paintOrder="fill markers stroke" />
                <circle id="arc6" className="circle" cx="150" cy="150" r="90" opacity=".49" fill="none" stroke="#632b26" strokeWidth="16" strokeLinecap="square" strokeOpacity=".99213" paintOrder="fill markers stroke" />
                <circle id="arc7" className="circle" cx="150" cy="150" r="90" opacity=".89" fill="none" stroke="#632b26" strokeWidth="8" strokeLinecap="square" strokeOpacity=".99213" paintOrder="fill markers stroke" />
                <circle id="arc8" className="circle" cx="150" cy="150" r="80" opacity=".79" fill="#4DD0E1"
                  fillOpacity="0" stroke="#632b26" strokeWidth="8" strokeLinecap="square" strokeOpacity=".99213" paintOrder="fill markers stroke" />
              </svg>
              <p className="loading-text">{status}</p>
            </div>
          ) : (
            <Form onSubmit={handleSubmit} className="upload-form">
              <Form.Group controlId="formMs6File">
                <Form.Label>Upload MS6 Excel File</Form.Label>
                <Form.Control type="file" name="ms6File" onChange={handleFileChange} />
              </Form.Group>
              <Form.Group controlId="formBmsFile">
                <Form.Label>Upload BMS Excel File</Form.Label>
                <Form.Control type="file" name="bmsFile" onChange={handleFileChange} />
              </Form.Group>
              <Button variant="primary" type="submit" className="generate-button">
                Generate Certificates
              </Button>
            </Form>
          )}
        </Col>
      </Row>
      <Modal show={showModal} onHide={handleCloseModal}>
        <Modal.Header closeButton>
          <Modal.Title>Success</Modal.Title>
        </Modal.Header>
        <Modal.Body>Certificate generated (check in the downloads folder of the device)</Modal.Body>
        <Modal.Footer>
          <Button variant="secondary" onClick={handleCloseModal}>
            Close
          </Button>
        </Modal.Footer>
      </Modal>
    </Container>
  );
}

export default App;
