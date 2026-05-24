import React, { useState, useEffect } from 'react';
import Header from './common/Header';
import Footer from './common/Footer';
import FloatingOrbs from './common/FloatingOrbs';
import { FORM_ACTION, APPS_SCRIPT_URL, ENTRY, jsonpRequest } from '../utils/api';

const RegisterScreen = ({ onNavigate, onShowToast }) => {
  // Form fields
  const [fullName, setFullName] = useState('');
  const [email, setEmail] = useState('');
  const [phone, setPhone] = useState('');
  const [selectedExam, setSelectedExam] = useState('');
  const [otherExam, setOtherExam] = useState('');

  // Field validation errors
  const [errors, setErrors] = useState({ name: false, email: false, phone: false });
  const [isLoading, setIsLoading] = useState(false);
  const [isSuccess, setIsSuccess] = useState(false);

  // Animated counters
  const [counters, setCounters] = useState({ pilots: 0, tests: 0, xp: 0 });

  useEffect(() => {
    // Animate counters
    const duration = 2000;
    const steps = duration / 16;
    let step = 0;

    const interval = setInterval(() => {
      step++;
      setCounters({
        pilots: Math.min(Math.round((124502 / steps) * step), 124502),
        tests: Math.min(Math.round((2400000 / steps) * step), 2400000),
        xp: Math.min(Math.round((1200000000 / steps) * step), 1200000000)
      });
      if (step >= steps) clearInterval(interval);
    }, 16);

    return () => clearInterval(interval);
  }, []);

  const formatCounter = (val, type) => {
    if (type === 'M') return (val / 1000).toFixed(1) + 'M';
    if (type === 'B') return (val / 1000000000).toFixed(1) + 'B';
    return val.toLocaleString();
  };

  const handlePhoneChange = (val) => {
    setPhone(val.replace(/[^0-9+\-\s()]/g, ''));
    setErrors(prev => ({ ...prev, phone: false }));
  };

  const validate = () => {
    let valid = true;
    const newErrors = { name: false, email: false, phone: false };

    if (!fullName.trim() || fullName.trim().length < 2) {
      newErrors.name = true;
      valid = false;
    }
    if (!email.trim() || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim())) {
      newErrors.email = true;
      valid = false;
    }
    if (!phone.trim() || !/^[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,6}$/.test(phone.trim())) {
      newErrors.phone = true;
      valid = false;
    }
    if (!selectedExam) {
      onShowToast('Please select your exam preparation category.', 'error');
      valid = false;
    }

    setErrors(newErrors);
    return valid;
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!validate()) return;

    setIsLoading(true);

    try {
      // Pre-registration check: Call Apps Script to verify email/phone uniqueness
      let checkRes = null;
      try {
        checkRes = await jsonpRequest(APPS_SCRIPT_URL, {
          action: 'checkRegistration',
          email: email.trim(),
          phone: phone.trim()
        });
      } catch (err) {
        console.warn('JSONP registration check failed, trying fetch fallback...', err);
        try {
          const url = `${APPS_SCRIPT_URL}?action=checkRegistration&email=${encodeURIComponent(email.trim())}&phone=${encodeURIComponent(phone.trim())}`;
          const res = await fetch(url, { redirect: 'follow' });
          const txt = await res.text();
          const cleaned = txt.trim().replace(/^[a-zA-Z_$][a-zA-Z0-9_$]*\s*\(/, '').replace(/\)\s*;?\s*$/, '');
          checkRes = JSON.parse(cleaned);
        } catch (fetchErr) {
          console.error('Fetch registration check also failed:', fetchErr);
        }
      }

      if (checkRes && checkRes.success === false) {
        onShowToast(checkRes.message, 'error');
        setIsLoading(false);
        return;
      }

      // Submit to Google Forms
      const isOther = selectedExam === 'Other';
      const customExamValue = otherExam.trim();

      const params = new URLSearchParams();
      params.append(ENTRY.fullName, fullName.trim());
      params.append(ENTRY.email, email.trim());
      params.append(ENTRY.phone, phone.trim());

      if (isOther && customExamValue) {
        params.append(ENTRY.preparation, '__other_option__');
        params.append(ENTRY.preparationOther, customExamValue);
      } else {
        params.append(ENTRY.preparation, selectedExam);
      }

      await fetch(FORM_ACTION, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: params.toString()
      });

      setIsSuccess(true);
      onShowToast(`Welcome aboard, ${fullName.split(' ')[0]}! 🚀`, 'success');

    } catch (err) {
      console.error('Submission error:', err);
      onShowToast('Something went wrong. Please try again.', 'error');
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="register-page-container">
      <FloatingOrbs />
      <Header mode="register" xpPoints={isSuccess ? 100 : 0} />

      <main>
        <div className="page-content">
          {/* Left Panel: Form */}
          <div className={`form-card ${isSuccess ? 'success-mode' : ''}`} id="formCard">
            {!isSuccess ? (
              <div className="form-body">
                <h1 className="form-title">Start Your Journey</h1>

                <div className="xp-banner">
                  <div className="xp-icon">⚡</div>
                  <div>Earn <strong>+100 XP</strong> points by completing your first test after registration.</div>
                </div>

                <form onSubmit={handleSubmit} noValidate>
                  <div className="field-group">
                    {/* Full Name */}
                    <div className={`field ${errors.name ? 'error' : ''}`}>
                      <input 
                        type="text" 
                        id="fullName" 
                        value={fullName}
                        onChange={(e) => { setFullName(e.target.value); setErrors(prev => ({ ...prev, name: false })); }}
                        placeholder=" " 
                        autoComplete="name" 
                      />
                      <label htmlFor="fullName">Full Name</label>
                      <div className="field-line"></div>
                      <div className="field-error">Please enter your full name</div>
                    </div>

                    {/* Email */}
                    <div className={`field ${errors.email ? 'error' : ''}`}>
                      <input 
                        type="email" 
                        id="emailAddress" 
                        value={email}
                        onChange={(e) => { setEmail(e.target.value); setErrors(prev => ({ ...prev, email: false })); }}
                        placeholder=" " 
                        autoComplete="email" 
                      />
                      <label htmlFor="emailAddress">Email Address</label>
                      <div className="field-line"></div>
                      <div className="field-error">Please enter a valid email address</div>
                    </div>

                    {/* Phone */}
                    <div className={`field ${errors.phone ? 'error' : ''}`}>
                      <input 
                        type="tel" 
                        id="phoneNumber" 
                        value={phone}
                        onChange={(e) => handlePhoneChange(e.target.value)}
                        placeholder=" " 
                        autoComplete="tel"
                        maxLength="15" 
                      />
                      <label htmlFor="phoneNumber">Phone Number</label>
                      <div className="field-line"></div>
                      <div className="field-error">Please enter a valid phone number</div>
                    </div>
                  </div>

                  <p className="prep-label">Preparation For</p>
                  <div className="prep-grid" role="group" aria-label="Select exam preparation">
                    {['UPSC', 'SSC', 'NEET', 'JEE', 'Banking', 'Other'].map(examName => (
                      <button
                        key={examName}
                        type="button"
                        className={`prep-btn ${selectedExam === examName ? 'selected' : ''}`}
                        onClick={() => setSelectedExam(examName)}
                      >
                        {examName}
                      </button>
                    ))}
                  </div>

                  {/* Other input option */}
                  <div className={`other-input-wrap ${selectedExam === 'Other' ? 'show' : ''}`}>
                    <div className="other-exam-field">
                      <input 
                        type="text" 
                        value={otherExam}
                        onChange={(e) => setOtherExam(e.target.value)}
                        placeholder="✏️ Type your exam name (e.g. GATE, CA, CDS...)" 
                        maxLength="60" 
                        autoComplete="off" 
                      />
                    </div>
                  </div>

                  <button type="submit" className="btn-register" disabled={isLoading}>
                    {!isLoading ? (
                      <span>Register &amp; Continue 🚀</span>
                    ) : (
                      <div className="spinner" style={{ display: 'block' }}></div>
                    )}
                  </button>
                </form>

                <p className="form-footer">
                  Already have an account? <a href="#" onClick={(e) => { e.preventDefault(); onNavigate('login'); }}>Login Pilot</a>
                </p>
              </div>
            ) : (
              <div className="success-overlay show">
                <div className="success-icon">✓</div>
                <div className="success-title">You're In, Pilot! 🎉</div>
                <p className="success-msg">Registration successful. Your journey begins now!</p>
                <div className="xp-earned">+100 XP Queued</div>
                <a 
                  href="#" 
                  onClick={(e) => { e.preventDefault(); onNavigate('login'); }}
                  style={{
                    display: 'inline-flex', alignItems: 'center', gap: '0.5rem',
                    marginTop: '0.8rem', padding: '0.75rem 2rem',
                    borderRadius: '12px', textDecoration: 'none',
                    background: 'linear-gradient(135deg,#4d8eff,#6fa3ff)',
                    color: '#fff', fontFamily: "'Sora',sans-serif",
                    fontSize: '0.95rem', fontWeight: 700,
                    boxShadow: '0 4px 18px rgba(77,142,255,0.4)',
                    transition: 'transform 0.2s,box-shadow 0.2s'
                  }}
                  onMouseOver={(e) => {
                    e.currentTarget.style.transform = 'translateY(-2px)';
                    e.currentTarget.style.boxShadow = '0 8px 28px rgba(77,142,255,0.5)';
                  }}
                  onMouseOut={(e) => {
                    e.currentTarget.style.transform = 'translateY(0)';
                    e.currentTarget.style.boxShadow = '0 4px 18px rgba(77,142,255,0.4)';
                  }}
                >
                  🚀 Login to Start Test
                </a>
              </div>
            )}
          </div>

          {/* Right Panel: Hero */}
          <div className="hero-panel">
            <img 
              className="hero-image" 
              src="https://images.unsplash.com/photo-1635070041078-e363dbe005cb?w=900&q=80"
              alt="Student mastering exams with futuristic holographic interface" 
              loading="lazy" 
            />
            <div className="hero-gradient"></div>

            <div className="live-indicator" aria-label="System Online">
              <div className="live-dot"></div>
              SYS. ONLINE
            </div>

            <div className="hero-xp-chip" aria-label="Potential 12k XP">
              <div className="xp-chip-icon">⭐</div>
              <div className="xp-chip-info">
                <span className="xp-chip-value">+12k</span>
                <span className="xp-chip-label">Potential XP</span>
              </div>
            </div>

            <div className="score-card">
              <div className="score-card-header">
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <rect x="3" y="3" width="18" height="18" rx="2" />
                  <path d="M3 9h18M9 21V9" />
                </svg>
                Target Score
              </div>
              <div className="score-value">98%</div>
              <div className="score-trend">↑ Current Trend: Ascending</div>
              <div className="score-bar">
                <div className="score-bar-fill"></div>
              </div>
            </div>

            <div className="hero-content">
              <h2>Master your Exams with Gamified Precision</h2>
              <p>Our platform turns every mock test into a mission. Gain experience, unlock elite ranks, and conquer competitive exams with our command-center dashboard.</p>
            </div>
          </div>
        </div>

        {/* Stats Bar */}
        <div className="stats-bar">
          <div className="stat-item">
            <span className="stat-label">Active Pilots</span>
            <span className="stat-value blue">{formatCounter(counters.pilots)}</span>
          </div>
          <div className="stat-divider"></div>
          <div className="stat-item">
            <span className="stat-label">Tests Completed</span>
            <span className="stat-value teal">{formatCounter(counters.tests, 'M')}</span>
          </div>
          <div className="stat-divider"></div>
          <div className="stat-item">
            <span className="stat-label">XP Generated</span>
            <span className="stat-value orange">{formatCounter(counters.xp, 'B')}</span>
          </div>
        </div>

        <Footer />
      </main>
    </div>
  );
};

export default RegisterScreen;
