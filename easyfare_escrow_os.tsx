import React, { useState, useEffect, useRef } from 'react';
import { 
  Search, Ticket, User, MapPin, Calendar, Clock, 
  CreditCard, QrCode, CheckCircle, ChevronLeft, Bus, 
  Settings, LogOut, ArrowRight, ShieldCheck, TrendingUp, Users,
  Sparkles, Send, Star, MessageSquare, Volume2, Image as ImageIcon, Reply,
  AlertTriangle, Shield, Check, RefreshCw, Smartphone, Globe, Lock, Info, Download, Share2,
  Package, Navigation, Compass, AlertCircle, PhoneCall, Briefcase, Plus, CheckSquare, Square,
  DollarSign, Building, ArrowUpRight, ArrowDownLeft, Wallet, Layers, ShieldAlert, Unlock, Zap, FileText, CheckCheck, Scale, FileCheck
} from 'lucide-react';

const STORAGE_KEY_TICKETS = 'easyfare_user_tickets_v9';
const STORAGE_KEY_PARCELS = 'easyfare_user_parcels_v9';
const STORAGE_KEY_AUTH = 'easyfare_auth_role_v9';
const STORAGE_KEY_USER = 'easyfare_logged_user_v9';
const STORAGE_KEY_REGISTERED_USERS = 'easyfare_registered_users_v9';
const STORAGE_KEY_WALLET = 'easyfare_station_wallet_v9';
const STORAGE_KEY_PLATFORM_WALLET = 'easyfare_platform_wallet_v10';
const STORAGE_KEY_TRIPS = 'easyfare_trips_v9';
const STORAGE_KEY_REFUNDS = 'easyfare_user_refunds_v9';
const STORAGE_KEY_OPERATOR_TERMS = 'easyfare_operator_terms_signed_v1';
const STORAGE_KEY_DISPUTES = 'easyfare_disputes_v1';
const STORAGE_KEY_MANIFEST = 'easyfare_immutable_manifest_ledger_v1';

const CITIES = ['Accra', 'Kumasi', 'Tamale', 'Takoradi', 'Cape Coast', 'Sunyani', 'Ho', 'Koforidua', 'Wa', 'Bolgatanga'];

const COMPANIES = [
  { id: 'c1', name: 'VIP Jeoun', rating: 4.8, color: 'bg-red-600', subaccountCode: 'ACCT_vipjeoun01', bank: 'GCB Bank', accountNumber: '1011020034921' },
  { id: 'c2', name: 'STC Ghana', rating: 4.6, color: 'bg-green-700', subaccountCode: 'ACCT_stcghana02', bank: 'Ecobank Ghana', accountNumber: '1442003892110' },
  { id: 'c3', name: 'O.A Travel', rating: 4.5, color: 'bg-blue-600', subaccountCode: 'ACCT_oatravel03', bank: 'Stanbic Bank', accountNumber: '9020014829102' },
];

const PREDEFINED_LUGGAGE_OPTIONS = [
  { id: 'suitcase', label: 'Standard Suitcase', desc: 'Hold luggage' },
  { id: 'ghana_must_go', label: '"Ghana Must Go" Bag', desc: 'Large woven bag' },
  { id: 'box', label: 'Carton / Sealed Box', desc: 'Packaged items' },
  { id: 'duffel', label: 'Duffel / Sports Bag', desc: 'Medium soft bag' },
  { id: 'foodstuff', label: 'Foodstuff / Agricultural Sacks', desc: 'Produce bag' },
  { id: 'fragile', label: 'Fragile / Electronic Equipment', desc: 'Handled with care' },
];

const initialTrips = [
  { id: 't1', companyId: 'c1', from: 'Accra', to: 'Kumasi', date: '2026-08-29', departureTime: '06:00 AM', arrivalTime: '10:30 AM', price: 150, totalSeats: 40, bookedSeats: [2, 3, 4, 15, 16], busType: 'Executive AC', status: 'En Route', currentLat: 6.12, currentLng: -0.78, escrowHeld: 0 },
  { id: 't2', companyId: 'c2', from: 'Accra', to: 'Kumasi', date: '2026-08-29', departureTime: '08:00 AM', arrivalTime: '12:30 PM', price: 120, totalSeats: 50, bookedSeats: [1, 2, 10], busType: 'Standard AC', status: 'Boarding', currentLat: 5.55, currentLng: -0.20, escrowHeld: 341.25 },
  { id: 't3', companyId: 'c3', from: 'Kumasi', to: 'Tamale', date: '2026-08-30', departureTime: '05:00 AM', arrivalTime: '02:00 PM', price: 200, totalSeats: 40, bookedSeats: [5, 6], busType: 'Executive AC', status: 'Scheduled', currentLat: 6.69, currentLng: -1.62, escrowHeld: 378.20 },
  { id: 't4', companyId: 'c1', from: 'Accra', to: 'Takoradi', date: '2026-08-29', departureTime: '07:30 AM', arrivalTime: '11:00 AM', price: 110, totalSeats: 40, bookedSeats: [], busType: 'Executive AC', status: 'Scheduled', currentLat: 5.55, currentLng: -0.20, escrowHeld: 0 }
];

const INITIAL_WALLET = { available: 14250.00, escrow: 719.45 };
const INITIAL_PLATFORM_WALLET = { balance: 4820.50, totalSplitCollected: 12450.00 };

const Toast = ({ message, type = 'success', onClose }) => {
  useEffect(() => {
    const timer = setTimeout(onClose, 4000);
    return () => clearTimeout(timer);
  }, [onClose]);

  return (
    <div className="fixed top-4 left-4 right-4 z-50 max-w-md mx-auto animate-bounce">
      <div className={`p-4 rounded-2xl shadow-xl flex items-center justify-between text-white ${type === 'success' ? 'bg-emerald-600' : type === 'warning' ? 'bg-amber-500' : 'bg-rose-600'}`}>
        <div className="flex items-center gap-3">
          {type === 'success' ? <CheckCircle size={22} /> : type === 'warning' ? <AlertTriangle size={22} /> : <AlertCircle size={22} />}
          <span className="font-medium text-sm">{message}</span>
        </div>
        <button onClick={onClose} className="text-white/80 hover:text-white font-bold px-2">✕</button>
      </div>
    </div>
  );
};

const Button = ({ children, onClick, variant = 'primary', className = '', fullWidth = false, icon: Icon, disabled = false }) => {
  const baseStyle = "flex items-center justify-center gap-2 px-4 py-3.5 font-semibold rounded-2xl transition-all duration-200 active:scale-95 disabled:opacity-50 disabled:cursor-not-allowed";
  const variants = {
    primary: "bg-indigo-600 text-white shadow-lg shadow-indigo-200 hover:bg-indigo-700",
    secondary: "bg-indigo-50 text-indigo-700 hover:bg-indigo-100",
    outline: "border-2 border-slate-200 text-slate-700 hover:border-slate-300 bg-white",
    danger: "bg-rose-50 text-rose-600 hover:bg-rose-100",
    success: "bg-emerald-600 text-white shadow-lg shadow-emerald-200 hover:bg-emerald-700",
    google: "bg-white border-2 border-slate-200 text-slate-800 hover:bg-slate-50 shadow-sm"
  };
  return (
    <button 
      onClick={onClick} 
      disabled={disabled}
      className={`${baseStyle} ${variants[variant]} ${fullWidth ? 'w-full' : ''} ${className}`}
    >
      {Icon && <Icon size={18} />}
      {children}
    </button>
  );
};

const Input = ({ label, icon: Icon, type = 'text', value, onChange, placeholder, error }) => (
  <div className="mb-4">
    {label && <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1.5">{label}</label>}
    <div className="relative">
      {Icon && <div className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400"><Icon size={20} /></div>}
      <input
        type={type}
        value={value}
        onChange={onChange}
        placeholder={placeholder}
        className={`w-full bg-slate-50 border text-slate-900 rounded-2xl px-4 py-3.5 focus:outline-none focus:ring-2 transition-all ${error ? 'border-rose-500 focus:ring-rose-500' : 'border-slate-200 focus:ring-indigo-500'} ${Icon ? 'pl-11' : ''}`}
      />
    </div>
    {error && <p className="text-xs text-rose-500 mt-1">{error}</p>}
  </div>
);

const AuthScreen = ({ onAuthSuccess, setToast }) => {
  const [mode, setMode] = useState('login');
  const [role, setRole] = useState('customer');
  const [identifier, setIdentifier] = useState('');
  const [password, setPassword] = useState('');
  const [name, setName] = useState('');
  const [phone, setPhone] = useState('');
  const [agreed, setAgreed] = useState(true);
  const [error, setError] = useState('');
  const [isGoogleLoading, setIsGoogleLoading] = useState(false);

  const handleGoogleAuth = () => {
    setIsGoogleLoading(true);
    setTimeout(() => {
      setIsGoogleLoading(false);
      const googleUser = { name: 'Kwame Mensah (Google)', email: 'kwame.mensah@gmail.com', phone: '+233 54 123 4567', role: 'customer' };
      try {
        const registered = JSON.parse(localStorage.getItem(STORAGE_KEY_REGISTERED_USERS) || '[]');
        if (!registered.some(u => u.email === googleUser.email)) {
          registered.push(googleUser);
          localStorage.setItem(STORAGE_KEY_REGISTERED_USERS, JSON.stringify(registered));
        }
      } catch (e) { console.error(e); }
      onAuthSuccess('customer', googleUser);
    }, 1200);
  };

  const handleSubmit = () => {
    setError('');
    if (!agreed) return setError('You must accept the Terms of Service to continue.');
    if (!identifier.trim()) return setError('Please enter your phone number or email.');
    if (!password.trim()) return setError('Please enter your password or PIN.');

    if (identifier === 'admin' || role === 'admin') {
      onAuthSuccess('admin', { name: 'Platform Admin', email: 'admin@easyfare.com', role: 'admin' });
      return;
    }

    if (mode === 'signup') {
      if (!name.trim()) return setError('Please enter your full name for registration.');
      const newUser = { name, email: identifier, phone: phone || identifier, role };
      try {
        const registered = JSON.parse(localStorage.getItem(STORAGE_KEY_REGISTERED_USERS) || '[]');
        if (registered.some(u => u.email === identifier)) return setError('An account with this email already exists. Please login.');
        registered.push(newUser);
        localStorage.setItem(STORAGE_KEY_REGISTERED_USERS, JSON.stringify(registered));
      } catch (e) { console.error(e); }
      onAuthSuccess(role, newUser);
    } else {
      try {
        const registered = JSON.parse(localStorage.getItem(STORAGE_KEY_REGISTERED_USERS) || '[]');
        const found = registered.find(u => u.email === identifier || u.phone === identifier);
        if (!found && role === 'customer') return setError('Account not found. Please sign up first.');
        const userObj = found || { name: identifier.includes('@') ? identifier.split('@')[0] : 'User', email: identifier, role };
        onAuthSuccess(role, userObj);
      } catch (e) {
        onAuthSuccess(role, { name: 'User', email: identifier, role });
      }
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6">
      <div className="w-full max-w-md bg-white rounded-[32px] shadow-xl shadow-slate-200/50 p-8 border border-slate-100">
        <div className="flex justify-center mb-4">
          <div className="bg-gradient-to-tr from-indigo-600 to-violet-600 p-4 rounded-3xl text-white shadow-lg shadow-indigo-200">
            <Bus size={36} />
          </div>
        </div>
        <h1 className="text-2xl font-black text-center text-slate-900 mb-1">EasyFare Ghana</h1>
        <p className="text-slate-500 text-center text-sm mb-6">Paystack Split Escrow OS</p>
        
        <div className="flex bg-slate-100 p-1.5 rounded-2xl mb-5">
          <button 
            className={`flex-1 py-2.5 rounded-xl font-bold text-xs tracking-wider uppercase transition-all ${mode === 'login' ? 'bg-white text-indigo-600 shadow-sm' : 'text-slate-500'}`}
            onClick={() => { setMode('login'); setError(''); }}
          >
            Sign In
          </button>
          <button 
            className={`flex-1 py-2.5 rounded-xl font-bold text-xs tracking-wider uppercase transition-all ${mode === 'signup' ? 'bg-white text-indigo-600 shadow-sm' : 'text-slate-500'}`}
            onClick={() => { setMode('signup'); setError(''); }}
          >
            Create Account
          </button>
        </div>

        <div className="flex bg-slate-50 border border-slate-200 p-1 rounded-xl mb-6">
          <button 
            className={`flex-1 py-2 rounded-lg font-bold text-xs transition-all ${role === 'customer' ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-600'}`}
            onClick={() => setRole('customer')}
          >
            Passenger
          </button>
          <button 
            className={`flex-1 py-2 rounded-lg font-bold text-xs transition-all ${role === 'admin' ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-600'}`}
            onClick={() => setRole('admin')}
          >
            Platform / Station Admin
          </button>
        </div>

        {error && (
          <div className="mb-4 p-3 bg-rose-50 border border-rose-200 text-rose-700 text-xs rounded-xl flex items-center gap-2">
            <AlertTriangle size={16} className="shrink-0" />
            <span>{error}</span>
          </div>
        )}

        {mode === 'signup' && (
          <Input label="Full Name" icon={User} placeholder="e.g., Kwame Mensah" value={name} onChange={e => setName(e.target.value)} />
        )}

        <Input label="Email or Phone Number" icon={User} placeholder="e.g., kwame@gmail.com or admin" value={identifier} onChange={e => setIdentifier(e.target.value)} />
        <Input label="Password / PIN" icon={ShieldCheck} type="password" placeholder="••••••••" value={password} onChange={e => setPassword(e.target.value)} />
        
        {role === 'customer' && mode === 'signup' && (
          <div className="mb-4">
            <Button variant="google" fullWidth onClick={handleGoogleAuth} disabled={isGoogleLoading}>
               {isGoogleLoading ? 'Connecting to Google...' : 'Sign up with Google'}
            </Button>
            <div className="flex items-center my-4">
              <div className="flex-1 border-t border-slate-200"></div>
              <span className="px-3 text-slate-400 text-xs uppercase font-bold">Or with email</span>
              <div className="flex-1 border-t border-slate-200"></div>
            </div>
          </div>
        )}

        <div className="flex items-start gap-3 mt-2 mb-6">
          <input 
            type="checkbox" 
            checked={agreed} 
            onChange={(e) => setAgreed(e.target.checked)} 
            className="mt-1 w-4 h-4 text-indigo-600 rounded border-slate-300 focus:ring-indigo-500"
          />
          <p className="text-xs text-slate-500 leading-relaxed">
            I agree to the <span className="text-indigo-600 font-semibold underline cursor-pointer">Terms of Service</span> and Paystack Split settlement protocols.
          </p>
        </div>

        <Button fullWidth onClick={handleSubmit}>
          {mode === 'signup' ? 'Complete Registration' : `Sign In as ${role === 'customer' ? 'Passenger' : 'Platform Admin'}`}
        </Button>
      </div>
    </div>
  );
};

const CustomerSearch = ({ onSearch, onSelectParcel, setToast }) => {
  const [from, setFrom] = useState('Accra');
  const [to, setTo] = useState('Kumasi');
  const [date, setDate] = useState('2026-08-29');
  const [magicQuery, setMagicQuery] = useState('');
  const [isMagicLoading, setIsMagicLoading] = useState(false);
  const [serviceType, setServiceType] = useState('bus');

  const handleMagicSearch = async () => {
    if (!magicQuery.trim()) return;
    setIsMagicLoading(true);
    try {
      const apiKey = "";
      const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${apiKey}`;
      const payload = {
          contents: [{ parts: [{ text: magicQuery }] }],
          generationConfig: {
              responseMimeType: "application/json",
              responseSchema: {
                  type: "OBJECT",
                  properties: {
                      from: { type: "STRING" },
                      to: { type: "STRING" },
                      date: { type: "STRING" }
                  }
              }
          },
          systemInstruction: {
            parts: [{ text: "Extract travel details. Available cities: Accra, Kumasi, Tamale, Takoradi, Cape Coast, Sunyani, Ho, Koforidua, Wa, Bolgatanga. Default origin Accra." }]
          }
      };

      const response = await fetch(apiUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
      const result = await response.json();
      if (result.candidates?.[0]?.content?.parts?.[0]?.text) {
          const data = JSON.parse(result.candidates[0].content.parts[0].text);
          if (data.from && CITIES.includes(data.from)) setFrom(data.from);
          if (data.to && CITIES.includes(data.to)) setTo(data.to);
          if (data.date) setDate(data.date);
          setMagicQuery('');
      }
    } catch (e) {
      const lower = magicQuery.toLowerCase();
      CITIES.forEach(c => {
        if (lower.includes(c.toLowerCase())) {
          if (from.toLowerCase() !== c.toLowerCase()) setTo(c);
        }
      });
    } finally {
      setIsMagicLoading(false);
    }
  };

  const executeSearch = () => {
    if (from === to) {
      setToast({ message: 'Origin and destination cannot be the same city.', type: 'error' });
      return;
    }
    if (serviceType === 'bus') {
      onSearch({ from, to, date });
    } else {
      onSelectParcel({ from, to, date });
    }
  };

  return (
    <div className="pb-32">
      <div className="bg-gradient-to-tr from-indigo-700 via-indigo-800 to-violet-800 pt-10 pb-28 px-6 rounded-b-[40px] text-white shadow-lg">
        <div className="flex justify-between items-center mb-4">
          <div>
            <span className="text-xs bg-white/20 px-3 py-1 rounded-full font-bold uppercase tracking-wider flex items-center gap-1 w-fit"><Zap size={12}/> Paystack Split OS</span>
            <h1 className="text-2xl font-black mt-2">Akwaaba! Where to?</h1>
          </div>
          <div className="w-10 h-10 bg-white/10 rounded-2xl flex items-center justify-center border border-white/20">
            <Globe size={20} />
          </div>
        </div>
        <p className="text-indigo-100 text-sm">Automated 5% platform fee split & 95% station escrow.</p>
      </div>

      <div className="px-6 -mt-20">
        <div className="bg-white rounded-[32px] p-6 shadow-xl shadow-slate-200/50 border border-slate-100">
          
          <div className="flex bg-slate-100 p-1.5 rounded-2xl mb-5">
            <button 
              className={`flex-1 py-2.5 rounded-xl font-bold text-xs tracking-wider uppercase flex items-center justify-center gap-2 transition-all ${serviceType === 'bus' ? 'bg-white text-indigo-600 shadow-sm' : 'text-slate-500'}`}
              onClick={() => setServiceType('bus')}
            >
              <Bus size={16} /> Bus Passenger
            </button>
            <button 
              className={`flex-1 py-2.5 rounded-xl font-bold text-xs tracking-wider uppercase flex items-center justify-center gap-2 transition-all ${serviceType === 'parcel' ? 'bg-white text-indigo-600 shadow-sm' : 'text-slate-500'}`}
              onClick={() => setServiceType('parcel')}
            >
              <Package size={16} /> Send Parcel
            </button>
          </div>

          <div className="mb-6 p-4 bg-indigo-50/80 rounded-2xl border border-indigo-100">
            <label className="flex items-center gap-2 text-xs font-bold uppercase tracking-wider text-indigo-900 mb-2">
              <Sparkles size={14} className="text-indigo-600" /> AI Prompt Search
            </label>
            <div className="flex gap-2">
              <input
                type="text"
                value={magicQuery}
                onChange={e => setMagicQuery(e.target.value)}
                placeholder="e.g., Book morning bus to Kumasi tomorrow"
                className="flex-1 bg-white border border-indigo-200 text-sm rounded-xl px-3.5 py-2.5 focus:outline-none focus:ring-2 focus:ring-indigo-500"
                onKeyPress={(e) => e.key === 'Enter' && handleMagicSearch()}
              />
              <button 
                onClick={handleMagicSearch}
                disabled={isMagicLoading || !magicQuery.trim()}
                className="bg-indigo-600 text-white p-2.5 rounded-xl hover:bg-indigo-700 disabled:opacity-50 transition-all flex items-center justify-center min-w-[44px]"
              >
                {isMagicLoading ? <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" /> : <Send size={18} />}
              </button>
            </div>
          </div>

          <div className="relative">
            <div className="mb-4">
              <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1.5">Origin City</label>
              <select 
                value={from} 
                onChange={e => setFrom(e.target.value)}
                className="w-full bg-slate-50 border border-slate-200 text-slate-900 rounded-2xl px-4 py-3.5 focus:outline-none focus:ring-2 focus:ring-indigo-500 font-medium"
              >
                {CITIES.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>

            <div className="absolute right-4 top-[45px] bg-white shadow-md p-2.5 rounded-full cursor-pointer hover:bg-slate-50 border border-slate-200 z-10"
                 onClick={() => { const temp = from; setFrom(to); setTo(temp); }}>
              <ArrowRight size={16} className="text-indigo-600 rotate-90" />
            </div>

            <div className="mb-4">
              <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1.5">Destination City</label>
              <select 
                value={to} 
                onChange={e => setTo(e.target.value)}
                className="w-full bg-slate-50 border border-slate-200 text-slate-900 rounded-2xl px-4 py-3.5 focus:outline-none focus:ring-2 focus:ring-indigo-500 font-medium"
              >
                {CITIES.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
          </div>

          <Input label="Date of Travel / Dispatch" icon={Calendar} type="date" value={date} onChange={e => setDate(e.target.value)} />
          
          <Button fullWidth className="mt-4" onClick={executeSearch}>
            {serviceType === 'bus' ? 'Search Express Buses' : 'Book Parcel Delivery Slot'}
          </Button>
        </div>
      </div>
    </div>
  );
};

const ParcelBooking = ({ details, onBack, onComplete, setToast }) => {
  const [receiverName, setReceiverName] = useState('');
  const [receiverPhone, setReceiverPhone] = useState('');
  const [weightKg, setWeightKg] = useState('5');
  const [description, setDescription] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);

  const parsedWeight = parseFloat(weightKg) || 0;
  const price = Math.max(30, parsedWeight * 8);

  const handleSubmit = () => {
    if (!receiverName.trim() || !receiverPhone.trim()) {
      setToast({ message: 'Please enter valid receiver details.', type: 'error' });
      return;
    }
    setIsSubmitting(true);
    setTimeout(() => {
      setIsSubmitting(false);
      onComplete({
        id: `PCL-${Math.floor(1000 + Math.random() * 9000)}`,
        from: details.from, to: details.to, receiverName, receiverPhone, weightKg: parsedWeight,
        description: description || 'General cargo', price, status: 'Dispatched', date: details.date,
        bookedAt: new Date().toISOString()
      });
    }, 1200);
  };

  return (
    <div className="pb-32 bg-slate-50 min-h-screen">
      <div className="bg-white px-4 py-4 flex items-center gap-4 shadow-sm border-b border-slate-100">
        <button onClick={onBack} className="p-2 hover:bg-slate-100 rounded-full"><ChevronLeft size={24} /></button>
        <div>
          <h2 className="font-bold text-slate-900">Parcel & Cargo Dispatch</h2>
          <p className="text-xs text-slate-500">{details.from} ➔ {details.to}</p>
        </div>
      </div>

      <div className="p-6 max-w-md mx-auto space-y-4">
        <div className="bg-gradient-to-tr from-indigo-700 to-violet-600 p-6 rounded-3xl text-white shadow-lg">
          <p className="text-indigo-200 text-xs font-bold uppercase tracking-wider mb-1">Estimated Cargo Fare</p>
          <h1 className="text-4xl font-black">GH₵{price}</h1>
          <p className="text-xs text-indigo-100 mt-2">Instant Paystack Split: 5% platform fee & 95% station escrow.</p>
        </div>

        <div className="bg-white p-6 rounded-3xl border border-slate-200 space-y-4">
          <h3 className="font-bold text-slate-900 text-sm uppercase tracking-wider">Receiver Details</h3>
          <Input label="Receiver Full Name" placeholder="e.g. Kwame Mensah" value={receiverName} onChange={e => setReceiverName(e.target.value)} />
          <Input label="Receiver MoMo Phone" placeholder="054 XXX XXXX" value={receiverPhone} onChange={e => setReceiverPhone(e.target.value)} />
          
          <h3 className="font-bold text-slate-900 text-sm uppercase tracking-wider pt-2">Package Weight & Info</h3>
          <Input label="Weight (Estimated KG)" type="number" value={weightKg} onChange={e => setWeightKg(e.target.value)} />
          <Input label="Item Description" placeholder="e.g., Foodstuff box, electronics" value={description} onChange={e => setDescription(e.target.value)} />
        </div>

        <Button fullWidth icon={Zap} onClick={handleSubmit} disabled={isSubmitting}>
          {isSubmitting ? 'Executing Paystack Split...' : `Pay GH₵${price} via Paystack Split`}
        </Button>
      </div>
    </div>
  );
};

const LiveBusTracker = ({ trip, onBack, setToast }) => {
  const [progress, setProgress] = useState(45);
  useEffect(() => {
    const interval = setInterval(() => setProgress(prev => (prev >= 95 ? 45 : prev + 2)), 3000);
    return () => clearInterval(interval);
  }, []);

  return (
    <div className="pb-32 bg-slate-50 min-h-screen">
      <div className="bg-white px-4 py-4 flex items-center gap-4 shadow-sm border-b border-slate-100">
        <button onClick={onBack} className="p-2 hover:bg-slate-100 rounded-full"><ChevronLeft size={24} /></button>
        <div>
          <h2 className="font-bold text-slate-900">Live GPS Tracker</h2>
          <p className="text-xs text-slate-500">{trip.from} to {trip.to}</p>
        </div>
      </div>

      <div className="p-6 max-w-md mx-auto space-y-4">
        <div className="bg-slate-900 rounded-3xl h-64 relative overflow-hidden flex flex-col items-center justify-center p-6 text-white text-center shadow-lg border border-slate-800">
          <div className="absolute inset-0 opacity-20 bg-[radial-gradient(#4f46e5_1px,transparent_1px)] [background-size:16px_16px]"></div>
          <div className="w-16 h-16 bg-indigo-600/90 rounded-full flex items-center justify-center border-4 border-white/20 animate-pulse mb-3 z-10 shadow-xl">
            <Bus size={32} className="text-white" />
          </div>
          <span className="bg-emerald-500/20 text-emerald-400 text-xs px-3 py-1 rounded-full font-bold uppercase tracking-wider z-10 mb-1 border border-emerald-500/30">
            Live • {trip.status} ({progress}%)
          </span>
          <p className="text-sm font-bold z-10">Passing Suhum Highway (N6)</p>
          <p className="text-xs text-slate-400 z-10">ETA Destination: {trip.arrivalTime}</p>
        </div>

        <div className="bg-white p-5 rounded-3xl border border-slate-200 space-y-3">
          <h3 className="font-bold text-slate-900 text-sm uppercase tracking-wider">Trip Updates & Escrow</h3>
          <div className="space-y-3 text-xs">
            <div className="flex gap-3">
              <div className="w-2 h-2 rounded-full bg-emerald-500 mt-1.5 shrink-0"></div>
              <div>
                <p className="font-bold text-slate-900">Paystack Split Verified</p>
                <p className="text-slate-500">5% platform creator fee credited instantly. Remaining 95% held in escrow for release upon departure.</p>
              </div>
            </div>
          </div>
        </div>

        <Button variant="outline" fullWidth icon={PhoneCall} onClick={() => setToast({ message: 'Dispatch desk is currently busy.', type: 'warning'})}>
          Contact Dispatch Desk
        </Button>
      </div>
    </div>
  );
};

const SearchResults = ({ searchParams, trips, onBack, onSelectTrip }) => {
  const results = trips.filter(t => {
    const matchFrom = !searchParams?.from || t.from.toLowerCase() === searchParams.from.toLowerCase();
    const matchTo = !searchParams?.to || t.to.toLowerCase() === searchParams.to.toLowerCase();
    return matchFrom && matchTo && t.status !== 'Canceled';
  });

  return (
    <div className="pb-32 bg-slate-50 min-h-screen">
      <div className="bg-white px-4 py-4 flex items-center gap-4 sticky top-0 z-20 shadow-sm border-b border-slate-100">
        <button onClick={onBack} className="p-2 hover:bg-slate-100 rounded-full"><ChevronLeft size={24} /></button>
        <div>
          <h2 className="font-bold text-slate-900">{searchParams?.from || 'Accra'} to {searchParams?.to || 'Kumasi'}</h2>
          <p className="text-xs text-slate-500">{searchParams?.date || 'Today'} • {results.length} departures found</p>
        </div>
      </div>

      <div className="p-4 space-y-4">
        {results.length === 0 ? (
          <div className="text-center py-16 bg-white rounded-3xl p-8 border border-slate-100 mt-4">
            <Bus size={48} className="mx-auto text-slate-300 mb-3" />
            <h3 className="font-bold text-slate-800 text-lg">No direct buses found</h3>
            <p className="text-xs text-slate-500 mt-1 mb-4">Try selecting another route combination.</p>
            <Button variant="outline" onClick={onBack}>Modify Search</Button>
          </div>
        ) : (
          results.map(trip => {
            const company = COMPANIES.find(c => c.id === trip.companyId) || COMPANIES[0];
            return (
              <div key={trip.id} onClick={() => onSelectTrip(trip)} className="bg-white p-5 rounded-3xl shadow-sm border border-slate-100 active:scale-[0.98] transition-transform cursor-pointer">
                <div className="flex justify-between items-center mb-4">
                  <div className="flex items-center gap-3">
                    <div className={`w-10 h-10 rounded-2xl ${company.color} flex items-center justify-center text-white font-bold text-xs shadow-md`}>
                      {company.name.substring(0, 2)}
                    </div>
                    <div>
                      <h3 className="font-bold text-slate-900">{company.name}</h3>
                      <p className="text-xs text-slate-500">{trip.busType} • ⭐ {company.rating}</p>
                    </div>
                  </div>
                  <div className="text-right">
                    <span className="text-lg font-black text-indigo-600">GH₵{trip.price}</span>
                  </div>
                </div>
                
                <div className="flex justify-between items-center relative bg-slate-50 p-4 rounded-2xl">
                  <div>
                    <p className="font-bold text-slate-800 text-sm">{trip.departureTime}</p>
                    <p className="text-xs text-slate-500">{trip.from}</p>
                  </div>
                  <div className="flex flex-col items-center px-4">
                    <div className="w-16 h-0.5 bg-slate-300 border-dashed border-t"></div>
                    <span className="text-[10px] text-slate-400 mt-1">Direct</span>
                  </div>
                  <div className="text-right">
                    <p className="font-bold text-slate-800 text-sm">{trip.arrivalTime}</p>
                    <p className="text-xs text-slate-500">{trip.to}</p>
                  </div>
                </div>
              </div>
            );
          })
        )}
      </div>
    </div>
  );
};

const SeatSelection = ({ trip, onBack, onContinue, setToast }) => {
  const [selectedSeats, setSelectedSeats] = useState([]);
  
  const toggleSeat = (seatNum) => {
    if (trip.bookedSeats.includes(seatNum)) return;
    if (selectedSeats.includes(seatNum)) {
      setSelectedSeats(selectedSeats.filter(s => s !== seatNum));
    } else {
      if (selectedSeats.length >= 6) {
        setToast({ message: 'Maximum 6 seats per booking.', type: 'warning' });
        return;
      }
      setSelectedSeats([...selectedSeats, seatNum]);
    }
  };

  const rows = Math.ceil(trip.totalSeats / 4);
  const layout = Array.from({ length: rows }, (_, rowIndex) => [
    rowIndex * 4 + 1, rowIndex * 4 + 2, 'aisle', rowIndex * 4 + 3, rowIndex * 4 + 4
  ]);

  return (
    <div className="pb-36 bg-slate-50 min-h-screen flex flex-col">
      <div className="bg-white px-4 py-4 flex items-center gap-4 shadow-sm z-10 border-b border-slate-100">
        <button onClick={onBack} className="p-2 hover:bg-slate-100 rounded-full"><ChevronLeft size={24} /></button>
        <div>
          <h2 className="font-bold text-slate-900">Select Seats</h2>
          <p className="text-xs text-slate-500">{COMPANIES.find(c=>c.id === trip.companyId)?.name} • {trip.departureTime}</p>
        </div>
      </div>

      <div className="flex-1 overflow-y-auto p-6 flex flex-col items-center">
        <div className="bg-white p-6 rounded-[36px] shadow-lg shadow-slate-200/50 border border-slate-200 w-full max-w-[280px]">
          <div className="flex justify-between items-center mb-6 pb-4 border-b border-slate-100">
             <span className="text-[10px] uppercase font-bold text-slate-400">Front / Driver</span>
             <div className="w-8 h-8 rounded-full bg-slate-100 flex items-center justify-center text-slate-400">
               <User size={16} />
             </div>
          </div>
          <div className="space-y-3">
            {layout.map((row, i) => (
              <div key={i} className="flex justify-between items-center">
                {row.map((seat, j) => {
                  if (seat === 'aisle') return <div key={`aisle-${i}`} className="w-6"></div>;
                  if (seat > trip.totalSeats) return <div key={`empty-${seat}`} className="w-9 h-9"></div>;
                  const isBooked = trip.bookedSeats.includes(seat);
                  const isSelected = selectedSeats.includes(seat);
                  return (
                    <button
                      key={seat} disabled={isBooked} onClick={() => toggleSeat(seat)}
                      className={`w-9 h-9 rounded-xl flex items-center justify-center font-bold text-xs transition-all ${isBooked ? 'bg-slate-100 text-slate-400 cursor-not-allowed' : isSelected ? 'bg-indigo-600 text-white shadow-md shadow-indigo-200 scale-105' : 'bg-white border-2 border-slate-200 text-slate-700 hover:border-indigo-400'}`}
                    >
                      {seat}
                    </button>
                  );
                })}
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="fixed bottom-0 left-0 right-0 bg-white p-5 shadow-[0_-10px_30px_rgba(0,0,0,0.08)] border-t border-slate-100 z-20">
        <div className="flex justify-between items-center mb-4 max-w-md mx-auto">
          <div>
            <p className="text-xs text-slate-500 uppercase font-bold">Total Fare</p>
            <p className="text-2xl font-black text-slate-900">GH₵{trip.price * selectedSeats.length}</p>
          </div>
        </div>
        <div className="max-w-md mx-auto">
          <Button fullWidth disabled={selectedSeats.length === 0} onClick={() => onContinue(selectedSeats)}>
            Continue to Paystack Split Checkout
          </Button>
        </div>
      </div>
    </div>
  );
};

const Checkout = ({ trip, seats, onBack, onComplete, setToast }) => {
  const [hasLuggage, setHasLuggage] = useState(false);
  const [selectedLuggageTypes, setSelectedLuggageTypes] = useState([]);
  const [customLuggage, setCustomLuggage] = useState('');
  const [method, setMethod] = useState('momo');
  const [phone, setPhone] = useState('0541234567');
  const [loading, setLoading] = useState(false);
  const [showBreakdownModal, setShowBreakdownModal] = useState(false);

  const grossTotal = trip.price * seats.length;
  const platformFee = parseFloat((grossTotal * 0.05).toFixed(2));
  const netCompanyPayout = parseFloat((grossTotal - platformFee).toFixed(2));
  const company = COMPANIES.find(c => c.id === trip.companyId) || COMPANIES[0];

  const handlePay = () => {
    if (method === 'momo' && !phone.trim()) {
      setToast({ message: 'Please enter a valid MoMo phone number.', type: 'error' });
      return;
    }

    const luggageSummary = hasLuggage ? {
      hasLuggage: true, types: selectedLuggageTypes.map(id => PREDEFINED_LUGGAGE_OPTIONS.find(o => o.id === id)?.label).filter(Boolean), customNote: customLuggage.trim()
    } : { hasLuggage: false, types: [], customNote: '' };

    setLoading(true);
    setTimeout(() => {
      setLoading(false);
      onComplete({
        luggage: luggageSummary,
        financials: { grossTotal, platformFee, netCompanyPayout, subaccount: company.subaccountCode }
      });
    }, 1500);
  };

  return (
    <div className="pb-36 bg-slate-50 min-h-screen">
       <div className="bg-white px-4 py-4 flex items-center gap-4 shadow-sm border-b border-slate-100">
        <button onClick={onBack} className="p-2 hover:bg-slate-100 rounded-full"><ChevronLeft size={24} /></button>
        <h2 className="font-bold text-slate-900">Paystack Split Checkout</h2>
      </div>

      <div className="p-6 max-w-md mx-auto space-y-6">
        <div className="bg-gradient-to-tr from-indigo-700 via-indigo-800 to-violet-800 rounded-3xl p-6 text-white shadow-lg relative overflow-hidden">
          <div className="flex justify-between items-start mb-2">
            <div>
              <p className="text-indigo-200 text-xs font-bold uppercase tracking-wider flex items-center gap-1.5"><Zap size={14}/> Paystack Split Total</p>
              <h1 className="text-4xl font-black">GH₵{grossTotal}</h1>
            </div>
            <button onClick={() => setShowBreakdownModal(true)} className="bg-white/20 hover:bg-white/30 text-xs font-bold px-3 py-1.5 rounded-full flex items-center gap-1 backdrop-blur-sm transition-all">
              <Layers size={14} /> Split Details
            </button>
          </div>
          <p className="text-xs text-indigo-100 mt-2">5% goes instantly to your creator account; 95% goes to station escrow.</p>
        </div>

        <div className="bg-white p-5 rounded-3xl border border-slate-200 space-y-4">
          <div className="flex justify-between items-center">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-indigo-50 text-indigo-600 rounded-2xl flex items-center justify-center shrink-0">
                <Briefcase size={20} />
              </div>
              <div>
                <h3 className="font-bold text-slate-900 text-sm">Hold Luggage</h3>
                <p className="text-xs text-slate-500">Bringing heavy bags?</p>
              </div>
            </div>
            <button 
              onClick={() => { setHasLuggage(!hasLuggage); if (hasLuggage) { setSelectedLuggageTypes([]); setCustomLuggage(''); } }}
              className={`px-4 py-2 rounded-xl text-xs font-bold uppercase transition-all ${hasLuggage ? 'bg-indigo-600 text-white shadow-md' : 'bg-slate-100 text-slate-600'}`}
            >
              {hasLuggage ? 'Yes' : 'No'}
            </button>
          </div>
        </div>

        <div>
          <h3 className="font-bold text-slate-900 mb-3 text-sm uppercase tracking-wider">Payment Method</h3>
          <div className="space-y-3 mb-6">
            <button onClick={() => setMethod('momo')} className={`w-full flex items-center justify-between p-4 rounded-2xl border-2 transition-all ${method === 'momo' ? 'border-indigo-600 bg-indigo-50/50' : 'border-slate-200 bg-white'}`}>
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-amber-100 text-amber-700 rounded-2xl flex items-center justify-center font-bold"><QrCode size={20} /></div>
                <div className="text-left">
                  <p className="font-bold text-slate-900 text-sm">Mobile Money (Paystack)</p>
                  <p className="text-xs text-slate-500">Automated split execution</p>
                </div>
              </div>
              {method === 'momo' && <CheckCircle className="text-indigo-600" size={20} />}
            </button>
          </div>

          {method === 'momo' && (
            <div className="bg-white p-5 rounded-3xl border border-slate-200 mb-6 space-y-4">
              <Input label="MoMo Phone Number" value={phone} onChange={e => setPhone(e.target.value)} placeholder="054 XXX XXXX" />
            </div>
          )}

          <Button fullWidth onClick={handlePay} disabled={loading} icon={Zap}>
            {loading ? 'Executing Paystack Split...' : `Pay GH₵{grossTotal} with Instant Split`}
          </Button>
        </div>
      </div>

      {showBreakdownModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-[32px] p-6 max-w-sm w-full space-y-4 shadow-2xl animate-fadeIn">
            <div className="flex justify-between items-center border-b border-slate-100 pb-3">
              <h3 className="font-black text-slate-900 flex items-center gap-2"><Zap size={18} className="text-indigo-600" /> Paystack Split Breakdown</h3>
              <button onClick={() => setShowBreakdownModal(false)} className="text-slate-400 hover:text-slate-600 font-bold">✕</button>
            </div>
            <div className="space-y-3 text-xs">
              <div className="flex justify-between py-1 border-b border-slate-100">
                <span className="text-slate-500">Gross Ticket Fare</span>
                <span className="font-bold text-slate-900">GH₵ {grossTotal.toFixed(2)}</span>
              </div>
              <div className="flex justify-between py-1 border-b border-slate-100 bg-emerald-50 px-2 rounded-lg">
                <span className="font-bold text-emerald-800">Your Platform Fee (5%)</span>
                <span className="font-black text-emerald-600">+ GH₵ {platformFee} (Instant)</span>
              </div>
              <div className="flex justify-between py-2 bg-indigo-50 px-3 rounded-xl border border-indigo-100">
                <span className="font-bold text-indigo-900 flex items-center gap-1"><Lock size={14}/> Station Escrow ({company.name})</span>
                <span className="font-black text-indigo-700">GH₵ {netCompanyPayout}</span>
              </div>
            </div>
            <p className="text-[10px] text-slate-500 italic text-center">Via Paystack Split Payments, the 5% platform fee goes straight to your creator bank account instantly without manual calculation.</p>
            <Button fullWidth onClick={() => setShowBreakdownModal(false)}>Understood</Button>
          </div>
        </div>
      )}
    </div>
  );
};

const LiveSecurityBadge = ({ date, time }) => {
  const [now, setNow] = useState(new Date());
  useEffect(() => {
    const timer = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);

  let diff = 0;
  try {
      const departureStr = `${date} ${time}`; 
      const departure = new Date(departureStr);
      diff = departure.getTime() - now.getTime();
  } catch (e) { diff = -1; }
  
  const formatTime = (ms) => {
    if (ms < 0) return "00:00:00";
    const h = Math.floor(ms / (1000 * 60 * 60));
    const m = Math.floor((ms % (1000 * 60 * 60)) / (1000 * 60));
    const s = Math.floor((ms % (1000 * 60)) / 1000);
    return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
  };

  const seconds = now.getSeconds();

  return (
    <div className="bg-slate-900/40 border-t border-white/10 px-6 py-3 flex items-center justify-between backdrop-blur-sm relative overflow-hidden">
       <div className="absolute top-0 bottom-0 w-32 bg-white/10 skew-x-12 blur-md transition-all duration-1000 ease-linear pointer-events-none" style={{ left: `${((seconds % 10) / 10) * 120}%`, transform: 'translateX(-100%)' }}></div>
       <div className="flex items-center gap-2 z-10">
          <div className="relative flex items-center justify-center w-3 h-3">
             <div className="absolute w-full h-full bg-emerald-400 rounded-full animate-ping opacity-75"></div>
             <div className="relative w-1.5 h-1.5 bg-emerald-500 rounded-full"></div>
          </div>
          <div>
            <p className="text-[10px] text-emerald-300 font-bold tracking-widest uppercase leading-none mb-0.5">Paystack Split</p>
            <p className="text-[8px] text-indigo-200 uppercase font-mono tracking-wider opacity-80 leading-none">5% Creator / 95% Escrow</p>
          </div>
       </div>
       <div className="text-right z-10 flex flex-col items-end">
         <div className="text-[11px] text-indigo-100 font-mono flex items-center gap-1.5 bg-black/20 px-2 py-1 rounded-lg border border-white/5">
           <ShieldCheck size={12} className="text-indigo-300" />
           {diff > 0 ? (
             <span>Departs in <span className="font-bold text-white tracking-widest">{formatTime(diff)}</span></span>
           ) : (
             <span className="text-rose-300 font-bold">Departed</span>
           )}
         </div>
       </div>
    </div>
  );
};

const MyTickets = ({ tickets, parcels, refunds, disputes, onTrackBus, onCancelTicket, onRaiseDispute, setToast }) => {
  const [activeTab, setActiveTab] = useState('tickets');
  const [confirmCancelIdx, setConfirmCancelIdx] = useState(null);
  const [showDisputeModal, setShowDisputeModal] = useState(false);
  const [disputeTicket, setDisputeTicket] = useState(null);
  const [disputeReason, setDisputeReason] = useState('Service failure / Bus breakdown / No-show');
  const [disputeDesc, setDisputeDesc] = useState('');

  const performCancel = () => {
    if (confirmCancelIdx !== null) {
      onCancelTicket(confirmCancelIdx);
      setConfirmCancelIdx(null);
    }
  };

  const handleOpenDispute = (ticket) => {
    setDisputeTicket(ticket);
    setShowDisputeModal(true);
  };

  const submitDispute = () => {
    if (!disputeDesc.trim()) {
      setToast({ message: 'Please describe the service issue for the chargeback shield.', type: 'error' });
      return;
    }
    onRaiseDispute({
      ticketId: disputeTicket.trip.id,
      route: `${disputeTicket.trip.from} to ${disputeTicket.trip.to}`,
      amount: disputeTicket.financials?.grossTotal || (disputeTicket.trip.price * disputeTicket.seats.length),
      reason: disputeReason,
      description: disputeDesc,
      date: new Date().toISOString(),
      status: 'Escalated to Operator'
    });
    setShowDisputeModal(false);
    setDisputeDesc('');
    setToast({ message: 'Dispute & Chargeback Shield invoked against operator!', type: 'success' });
  };

  return (
    <div className="pb-32 bg-slate-50 min-h-screen">
      <div className="bg-white px-6 py-6 shadow-sm border-b border-slate-100">
        <h1 className="text-2xl font-black text-slate-900">My Orders & Passes</h1>
        <div className="flex bg-slate-100 p-1 rounded-2xl mt-4">
          <button className={`flex-1 py-2 rounded-xl text-xs font-bold uppercase ${activeTab === 'tickets' ? 'bg-white text-indigo-600 shadow-sm' : 'text-slate-500'}`} onClick={() => setActiveTab('tickets')}>
            Tickets ({tickets.length})
          </button>
          <button className={`flex-1 py-2 rounded-xl text-xs font-bold uppercase ${activeTab === 'parcels' ? 'bg-white text-indigo-600 shadow-sm' : 'text-slate-500'}`} onClick={() => setActiveTab('parcels')}>
            Parcels ({parcels.length})
          </button>
          <button className={`flex-1 py-2 rounded-xl text-xs font-bold uppercase ${activeTab === 'disputes' ? 'bg-white text-indigo-600 shadow-sm' : 'text-slate-500'}`} onClick={() => setActiveTab('disputes')}>
            Disputes ({disputes.length})
          </button>
        </div>
      </div>

      {activeTab === 'tickets' && tickets.length === 0 && (
          <div className="flex flex-col items-center justify-center min-h-[50vh] p-6 text-center">
            <Ticket size={48} className="text-slate-300 mb-3" />
            <h3 className="font-bold text-slate-800">No Bus Passes Found</h3>
            <p className="text-xs text-slate-500 max-w-xs mt-1">Book your bus tickets with secure Paystack split protection.</p>
          </div>
      )}

      {activeTab === 'tickets' && tickets.length > 0 && (
          <div className="p-4 space-y-6 max-w-md mx-auto">
            {tickets.map((ticket, idx) => (
              <div key={idx} className="bg-white rounded-[32px] overflow-hidden shadow-xl shadow-slate-200/50 border border-slate-100 relative">
                <div className="p-0 bg-gradient-to-tr from-indigo-700 via-indigo-800 to-violet-800 text-white">
                  <div className="p-6 pb-5">
                    <div className="flex justify-between items-center mb-4">
                      <span className="bg-white/20 px-3 py-1 rounded-full text-[10px] font-bold tracking-widest uppercase flex items-center gap-1">
                        <Zap size={10} /> Split Executed
                      </span>
                      <div className="flex gap-2">
                        <button onClick={() => onTrackBus(ticket.trip)} className="bg-emerald-500 text-white text-xs font-bold px-3 py-1 rounded-full flex items-center gap-1 shadow-sm hover:bg-emerald-600">
                          <Navigation size={12} /> Tracker
                        </button>
                        <button onClick={() => setConfirmCancelIdx(idx)} className="bg-rose-500/80 hover:bg-rose-600 text-white text-xs font-bold px-2.5 py-1 rounded-full transition-all" title="Cancel Ticket">✕</button>
                      </div>
                    </div>
                    <div className="flex justify-between items-end">
                      <div><p className="text-indigo-200 text-xs uppercase font-bold">From</p><p className="text-xl font-black">{ticket.trip.from}</p></div>
                      <ArrowRight className="text-indigo-300 mx-2 mb-2" />
                      <div className="text-right"><p className="text-indigo-200 text-xs uppercase font-bold">To</p><p className="text-xl font-black">{ticket.trip.to}</p></div>
                    </div>
                  </div>
                  <LiveSecurityBadge date={ticket.trip.date} time={ticket.trip.departureTime} />
                </div>
                <div className="p-6 bg-white space-y-4">
                  <div className="flex justify-between items-center">
                    <div>
                      <p className="text-[10px] uppercase font-bold text-slate-400">Assigned Seats</p>
                      <p className="font-black text-slate-900 text-xl">{ticket.seats.join(', ')}</p>
                      {ticket.boardingVerified && (
                        <p className="text-[10px] text-emerald-600 font-bold uppercase mt-1 flex items-center gap-1">
                          <CheckCircle size={12} /> Scanned & Logged on Manifest Ledger
                        </p>
                      )}
                    </div>
                    <Button variant="danger" icon={ShieldAlert} onClick={() => handleOpenDispute(ticket)}>
                      Dispute & Shield
                    </Button>
                  </div>
                </div>
              </div>
            ))}
          </div>
      )}

      {activeTab === 'disputes' && (
        <div className="p-4 space-y-4 max-w-md mx-auto">
          {disputes.length === 0 ? (
            <div className="text-center py-16 bg-white rounded-3xl p-8 border border-slate-100 mt-4">
              <ShieldCheck size={48} className="mx-auto text-indigo-300 mb-3" />
              <h3 className="font-bold text-slate-800 text-lg">No Active Disputes</h3>
              <p className="text-xs text-slate-500 mt-1">Chargeback Shield protects your platform by shifting liability directly to the operator.</p>
            </div>
          ) : (
            disputes.map((dis, idx) => (
              <div key={idx} className="bg-white p-5 rounded-3xl shadow-sm border border-slate-200 space-y-2">
                <div className="flex justify-between items-center">
                  <span className="text-xs font-bold uppercase bg-amber-100 text-amber-800 px-2.5 py-1 rounded-full">{dis.status}</span>
                  <span className="text-xs text-slate-400">{new Date(dis.date).toLocaleDateString()}</span>
                </div>
                <h4 className="font-bold text-slate-900">{dis.route}</h4>
                <p className="text-xs text-slate-600"><b>Reason:</b> {dis.reason}</p>
                <p className="text-xs text-slate-500 italic">"{dis.description}"</p>
                <div className="flex justify-between items-center pt-2 border-t border-slate-100 text-xs">
                  <span className="text-slate-500">Operator Liability Held:</span>
                  <span className="font-black text-rose-600">GH₵ {dis.amount.toFixed(2)}</span>
                </div>
              </div>
            ))
          )}
        </div>
      )}

      {showDisputeModal && disputeTicket && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-[32px] p-6 max-w-md w-full space-y-4 shadow-2xl animate-fadeIn">
            <div className="flex justify-between items-center border-b border-slate-100 pb-3">
              <h3 className="font-black text-slate-900 flex items-center gap-2"><ShieldAlert size={18} className="text-rose-600" /> Dispute & Chargeback Shield</h3>
              <button onClick={() => setShowDisputeModal(false)} className="text-slate-400 hover:text-slate-600 font-bold">✕</button>
            </div>
            <p className="text-xs text-slate-600 leading-relaxed">
              Under your digital operator terms agreement, transport providers are fully liable for bad service and chargeback disputes. Invoking this shifts financial liability away from your platform.
            </p>
            
            <div className="space-y-3">
              <div>
                <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">Issue Category</label>
                <select 
                  value={disputeReason} 
                  onChange={e => setDisputeReason(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 text-slate-900 rounded-xl px-3 py-2.5 text-xs font-medium"
                >
                  <option value="Service failure / Bus breakdown">Service failure / Bus breakdown</option>
                  <option value="Driver misconduct / No show">Driver misconduct / No show</option>
                  <option value="Unsafe transit conditions">Unsafe transit conditions</option>
                  <option value="Luggage damage / Loss">Luggage damage / Loss</option>
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-1">Detailed Description</label>
                <textarea 
                  rows="3"
                  value={disputeDesc}
                  onChange={e => setDisputeDesc(e.target.value)}
                  placeholder="Describe the bad service or operator failure..."
                  className="w-full bg-slate-50 border border-slate-200 text-slate-900 rounded-xl p-3 text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500"
                ></textarea>
              </div>
            </div>

            <div className="flex gap-2 pt-2">
              <Button variant="outline" fullWidth onClick={() => setShowDisputeModal(false)}>Cancel</Button>
              <Button variant="danger" fullWidth onClick={submitDispute}>Shield & File Dispute</Button>
            </div>
          </div>
        </div>
      )}

      {confirmCancelIdx !== null && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4 animate-fadeIn">
           <div className="bg-white rounded-3xl p-6 w-full max-w-sm text-center space-y-4">
              <div className="w-16 h-16 bg-rose-100 text-rose-600 rounded-full flex items-center justify-center mx-auto mb-2">
                <AlertTriangle size={28} />
              </div>
              <h3 className="font-black text-slate-900 text-lg">Cancel Ticket?</h3>
              <p className="text-sm text-slate-500 leading-relaxed">This action will release your seats and trigger an automatic refund from the Escrow hold to your account.</p>
              <div className="flex gap-3 pt-2">
                <Button variant="outline" fullWidth onClick={() => setConfirmCancelIdx(null)}>Keep Ticket</Button>
                <Button variant="danger" fullWidth onClick={performCancel}>Yes, Cancel</Button>
              </div>
           </div>
        </div>
      )}
    </div>
  );
};

const AdminDashboard = ({ stationWallet, setStationWallet, platformWallet, setPlatformWallet, trips, setTrips, refunds, setRefunds, disputes, setDisputes, myTickets, setMyTickets, manifestLedger, setManifestLedger, onLogout, setToast }) => {
  const [activeTab, setActiveTab] = useState('creator');
  const [isSignedTerms, setIsSignedTerms] = useState(() => localStorage.getItem(STORAGE_KEY_OPERATOR_TERMS) === 'true');
  const [showSignModal, setShowSignModal] = useState(false);
  const [operatorRepName, setOperatorRepName] = useState('Kofi Annan (Operations Director)');
  const [operatorIdNumber, setOperatorIdNumber] = useState('GHA-892104921-1');
  const [scanInputCode, setScanInputCode] = useState('');

  const handleSignTerms = () => {
    if (!operatorRepName.trim() || !operatorIdNumber.trim()) {
      setToast({ message: 'Please enter authorized representative name and ID.', type: 'error' });
      return;
    }
    setIsSignedTerms(true);
    localStorage.setItem(STORAGE_KEY_OPERATOR_TERMS, 'true');
    setShowSignModal(false);
    setToast({ message: 'Digital Operator Terms successfully signed! Platform shielded from direct chargebacks.', type: 'success' });
  };

  const handleWithdrawCreator = () => {
    if (platformWallet.balance <= 0) {
      setToast({ message: 'No platform split earnings to withdraw.', type: 'warning' });
      return;
    }
    setToast({ message: `Successfully transferred GH₵ ${platformWallet.balance.toFixed(2)} directly to your Creator Bank Account!`, type: 'success' });
    setPlatformWallet(prev => ({ ...prev, balance: 0 }));
  };

  const handleWithdrawStation = () => {
    if (!isSignedTerms) {
      setToast({ message: 'Operator must sign the Dispute & Chargeback Shield terms before station withdrawals!', type: 'warning' });
      setShowSignModal(true);
      return;
    }
    if (stationWallet.available <= 0) {
      setToast({ message: 'No available funds to withdraw yet.', type: 'warning' });
      return;
    }
    setToast({ message: `Successfully requested payout of GH₵ ${stationWallet.available.toFixed(2)} to corporate station bank account.`, type: 'success' });
    setStationWallet(prev => ({ ...prev, available: 0 }));
  };

  const handleDispatchBus = (tripId) => {
    const trip = trips.find(t => t.id === tripId);
    if(!trip) return;
    
    const releasedAmount = trip.escrowHeld || 0;
    
    const updatedTrips = trips.map(t =>
      t.id === tripId ? { ...t, status: 'En Route', escrowHeld: 0 } : t
    );
    setTrips(updatedTrips);

    setStationWallet(prev => ({
      available: prev.available + releasedAmount,
      escrow: prev.escrow - releasedAmount
    }));

    setToast({ message: `Departure confirmed! GH₵ ${releasedAmount.toFixed(2)} released from Escrow to Available Balance.`, type: 'success' });
  };

  const handleCancelTripByOperator = (tripId) => {
    const trip = trips.find(t => t.id === tripId);
    if (!trip) return;

    const escrowHeld = trip.escrowHeld || (trip.price * trip.bookedSeats.length * 0.95);
    
    const affectedTickets = myTickets.filter(t => t.trip.id === tripId);
    let totalRefundedToPassengers = 0;

    affectedTickets.forEach(ticket => {
      const gross = ticket.financials?.grossTotal || (trip.price * ticket.seats.length);
      const refAmt = gross * 0.90;
      totalRefundedToPassengers += refAmt;

      setRefunds(prev => [{
        route: `${trip.from} to ${trip.to}`,
        grossTotal: gross,
        refundAmount: refAmt,
        date: new Date().toISOString()
      }, ...prev]);
    });

    const updatedTrips = trips.map(t =>
      t.id === tripId ? { ...t, status: 'Canceled', escrowHeld: 0 } : t
    );
    setTrips(updatedTrips);

    const remainingTickets = myTickets.filter(t => t.trip.id !== tripId);
    setMyTickets(remainingTickets);

    setStationWallet(prev => ({
      ...prev,
      escrow: Math.max(0, prev.escrow - escrowHeld)
    }));

    setToast({ message: `Trip canceled by operator. Automated 90% refunds (GH₵ ${totalRefundedToPassengers.toFixed(2)}) sent to passengers!`, type: 'success' });
  };

  const handleScanTicket = (ticketCodeOrSeat) => {
    const query = (ticketCodeOrSeat || scanInputCode).trim();
    if (!query) return;

    // Find ticket in booked tickets or create manifest record
    const ticketIdx = myTickets.findIndex(t => t.seats.map(s => s.toString()).includes(query) || t.trip.id.toLowerCase() === query.toLowerCase());
    
    const timestamp = new Date().toISOString();
    const hashSignature = `HASH-${Math.random().toString(36).substring(2, 10).toUpperCase()}-${Date.now().toString(36)}`;
    
    const newManifestEntry = {
      id: `MAN-${Math.floor(1000 + Math.random() * 9000)}`,
      query,
      scannedAt: timestamp,
      hashSignature,
      stationVerified: true,
      operatorBlockedClaim: "Immune Ledger: Customer presence verified cryptographically. Commission locked & secure."
    };

    setManifestLedger(prev => [newManifestEntry, ...prev]);

    if (ticketIdx !== -1) {
      const updatedTickets = [...myTickets];
      updatedTickets[ticketIdx] = { ...updatedTickets[ticketIdx], boardingVerified: true };
      setMyTickets(updatedTickets);
    }

    setScanInputCode('');
    setToast({ message: `Ticket scanned & permanently recorded on Immutable Ledger! (${hashSignature})`, type: 'success' });
  };

  return (
    <div className="pb-32 bg-slate-50 min-h-screen">
      <div className="bg-slate-900 pt-8 pb-6 px-6 text-white sticky top-0 z-20 shadow-lg">
        <div className="flex justify-between items-center mb-6">
          <div>
            <span className="text-[10px] bg-emerald-500/30 text-emerald-300 px-2.5 py-1 rounded-full font-bold uppercase tracking-wider flex items-center gap-1 w-fit"><Zap size={10}/> Paystack Split OS</span>
            <h1 className="text-2xl font-black mt-1">Admin & Creator Dashboard</h1>
          </div>
          <button onClick={onLogout} className="bg-slate-800 p-2.5 rounded-2xl hover:bg-slate-700 text-slate-300">
            <LogOut size={20} />
          </button>
        </div>
        
        <div className="flex bg-slate-800 p-1 rounded-2xl overflow-x-auto gap-1">
          {[
            { id: 'creator', label: '🛡️ Creator' },
            { id: 'manifest', label: '📜 Ledger' },
            { id: 'shield', label: '⚖️ Shield' },
            { id: 'station', label: 'Escrow' },
            { id: 'dispatch', label: 'Departures' }
          ].map(tab => (
            <button 
              key={tab.id}
              className={`flex-1 py-2.5 px-3 rounded-xl font-bold text-[11px] uppercase tracking-wider whitespace-nowrap transition-all ${activeTab === tab.id ? 'bg-indigo-600 text-white shadow-md' : 'text-slate-400'}`}
              onClick={() => setActiveTab(tab.id)}
            >
              {tab.label}
            </button>
          ))}
        </div>
      </div>

      {activeTab === 'creator' && (
        <div className="p-6 space-y-6 max-w-md mx-auto">
          <div className="bg-gradient-to-br from-emerald-900 via-emerald-950 to-slate-900 rounded-[32px] p-6 text-white shadow-xl relative overflow-hidden border border-emerald-500/30">
            <div className="flex justify-between items-center mb-4">
              <span className="text-xs text-emerald-400 font-bold uppercase tracking-wider flex items-center gap-1.5"><ShieldCheck size={16} /> App Creator Protection</span>
              <span className="bg-emerald-500/20 text-emerald-400 text-[10px] px-2.5 py-1 rounded-full font-bold uppercase border border-emerald-500/30">Auto 5% Split</span>
            </div>
            <p className="text-xs text-emerald-200 uppercase font-bold mb-1">Your Instant Split Balance</p>
            <h1 className="text-4xl font-black mb-2">GH₵ {platformWallet.balance.toLocaleString('en-US', { minimumFractionDigits: 2 })}</h1>
            <p className="text-[11px] text-emerald-100/70 mb-5">Total split fees collected automatically across all customer ticket checkouts: GH₵ {platformWallet.totalSplitCollected.toFixed(2)}</p>
            
            <div className="pt-2 border-t border-emerald-500/30">
              <Button fullWidth variant="success" icon={ArrowUpRight} onClick={handleWithdrawCreator} disabled={platformWallet.balance <= 0}>
                 Withdraw to Creator Bank Account
              </Button>
            </div>
          </div>

          <div className="bg-white p-5 rounded-3xl border border-slate-200 space-y-3">
            <h3 className="font-bold text-slate-900 text-sm uppercase tracking-wider flex items-center gap-2"><Zap size={16} className="text-indigo-600" /> How Paystack Split Works</h3>
            <p className="text-xs text-slate-600 leading-relaxed">
              Whenever a passenger purchases a ticket (e.g. GH₵ 150), Paystack Split Payments automatically routes <b>5%</b> directly to your creator bank account instantly, while the remaining <b>95%</b> goes straight to the station's escrow balance. You never have to manually calculate or distribute funds!
            </p>
          </div>
        </div>
      )}

      {activeTab === 'manifest' && (
        <div className="p-6 space-y-6 max-w-md mx-auto">
          <div className="bg-gradient-to-br from-indigo-900 via-indigo-950 to-slate-900 rounded-[32px] p-6 text-white shadow-xl relative overflow-hidden border border-indigo-500/30 space-y-4">
            <div className="flex justify-between items-center">
              <span className="text-xs text-indigo-300 font-bold uppercase tracking-wider flex items-center gap-1.5"><FileCheck size={16} /> Immutable Manifest Ledger</span>
              <span className="bg-emerald-500/20 text-emerald-400 text-[10px] px-2.5 py-1 rounded-full font-bold uppercase border border-emerald-500/30">Anti-Fraud Active</span>
            </div>
            <h2 className="text-xl font-black">Scan & Log Passenger Presence</h2>
            <p className="text-xs text-indigo-100 leading-relaxed">
              Every ticket scanned at the station terminal is permanently hashed and logged to prevent stations from falsely claiming a customer "never showed up" to evade commission payouts.
            </p>
            
            <div className="space-y-3 pt-2">
              <Input label="Enter Seat No. or Ticket ID to Scan" placeholder="e.g. 2 or t1" value={scanInputCode} onChange={e => setScanInputCode(e.target.value)} />
              <Button fullWidth icon={QrCode} onClick={() => handleScanTicket()}>
                Record Scan on Immutable Ledger
              </Button>
            </div>
          </div>

          <div className="bg-white p-5 rounded-3xl border border-slate-200 space-y-3">
            <h3 className="font-bold text-slate-900 text-sm uppercase tracking-wider">Permanent Ledger Entries ({manifestLedger.length})</h3>
            {manifestLedger.length === 0 ? (
              <p className="text-xs text-slate-500">No boarding passes scanned yet. Scan a ticket above to write to the ledger.</p>
            ) : (
              manifestLedger.map((item, idx) => (
                <div key={idx} className="p-4 bg-slate-50 rounded-2xl border border-slate-200 text-xs space-y-2">
                  <div className="flex justify-between items-center font-bold text-slate-900">
                    <span className="flex items-center gap-1 text-emerald-700"><CheckCircle size={14} /> Scan ID: {item.id}</span>
                    <span className="font-mono text-[10px] bg-indigo-50 text-indigo-700 px-2 py-0.5 rounded">{item.hashSignature}</span>
                  </div>
                  <p className="text-slate-600"><b>Searched/Scanned Query:</b> Seat/ID #{item.query}</p>
                  <p className="text-slate-400 text-[10px]">{new Date(item.scannedAt).toLocaleString()}</p>
                  <div className="bg-emerald-50 text-emerald-800 p-2 rounded-xl text-[10px] font-medium border border-emerald-100">
                    {item.operatorBlockedClaim}
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      )}

      {activeTab === 'shield' && (
        <div className="p-6 space-y-6 max-w-md mx-auto">
          <div className="bg-gradient-to-br from-indigo-900 via-indigo-950 to-slate-900 rounded-[32px] p-6 text-white shadow-xl relative overflow-hidden border border-indigo-500/30 space-y-4">
            <div className="flex justify-between items-center">
              <span className="text-xs text-indigo-300 font-bold uppercase tracking-wider flex items-center gap-1.5"><Scale size={16} /> Operator Liability Agreement</span>
              <span className={`text-[10px] px-3 py-1 rounded-full font-bold uppercase ${isSignedTerms ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-rose-500/20 text-rose-300 border border-rose-500/30'}`}>
                {isSignedTerms ? 'Digitally Signed & Active' : 'Action Required'}
              </span>
            </div>
            
            <h2 className="text-xl font-black">Dispute & Chargeback Shield</h2>
            
            <p className="text-xs text-indigo-100 leading-relaxed">
              Transport operators must sign digital terms holding them strictly liable for bad service, bus breakdowns, and passenger chargebacks. This safeguards your software platform from direct financial liability.
            </p>

            {isSignedTerms ? (
              <div className="bg-white/10 p-4 rounded-2xl border border-white/10 space-y-2 text-xs">
                <div className="flex items-center gap-2 text-emerald-300 font-bold">
                  <CheckCheck size={16} /> Agreement Validated & Secured
                </div>
                <p className="text-indigo-200">Authorized Rep: <b>{operatorRepName}</b></p>
                <p className="text-indigo-200">ID / National Card: <b>{operatorIdNumber}</b></p>
                <p className="text-[10px] text-indigo-300 italic pt-1">All incoming disputes and customer chargebacks are legally bound to operator escrow deductions.</p>
              </div>
            ) : (
              <Button fullWidth variant="danger" icon={FileText} onClick={() => setShowSignModal(true)}>
                Sign Operator Digital Terms Now
              </Button>
            )}
          </div>

          <div className="bg-white p-5 rounded-3xl border border-slate-200 space-y-3">
            <h3 className="font-bold text-slate-900 text-sm uppercase tracking-wider">Active Disputes Shield ({disputes.length})</h3>
            {disputes.length === 0 ? (
              <p className="text-xs text-slate-500">No active chargeback disputes currently filed.</p>
            ) : (
              disputes.map((d, i) => (
                <div key={i} className="p-3 bg-slate-50 rounded-2xl border border-slate-200 text-xs space-y-1">
                  <div className="flex justify-between font-bold text-slate-900">
                    <span>{d.route}</span>
                    <span className="text-rose-600">GH₵ {d.amount.toFixed(2)}</span>
                  </div>
                  <p className="text-slate-500">{d.reason}</p>
                </div>
              ))
            )}
          </div>
        </div>
      )}

      {activeTab === 'station' && (
        <div className="p-6 space-y-6 max-w-md mx-auto">
          <div className="grid grid-cols-1 gap-4 mb-2">
            <div className="bg-gradient-to-br from-indigo-900 via-indigo-950 to-slate-900 rounded-[32px] p-6 text-white shadow-xl relative overflow-hidden border border-indigo-500/30">
              <div className="flex justify-between items-center mb-4">
                <span className="text-xs text-emerald-400 font-bold uppercase tracking-wider flex items-center gap-1.5"><Wallet size={16} /> Station Available Payout</span>
                <span className="bg-emerald-500/20 text-emerald-400 text-[10px] px-2.5 py-1 rounded-full font-bold uppercase border border-emerald-500/30">Cleared</span>
              </div>
              <h1 className="text-4xl font-black mb-5">GH₵ {stationWallet.available.toLocaleString('en-US', { minimumFractionDigits: 2 })}</h1>
              
              <div className="pt-2 border-t border-indigo-500/30">
                <Button fullWidth variant="success" icon={ArrowUpRight} onClick={handleWithdrawStation} disabled={stationWallet.available <= 0}>
                   Withdraw to Station Bank
                </Button>
              </div>
            </div>

            <div className="bg-white rounded-[32px] p-6 shadow-sm border border-slate-200 relative overflow-hidden">
               <div className="absolute top-0 right-0 p-4 opacity-5 pointer-events-none">
                 <Lock size={120} />
               </div>
               <div className="relative z-10">
                 <p className="text-xs text-amber-600 font-bold uppercase flex items-center gap-1.5 mb-2"><Lock size={16}/> Locked in Escrow (95%)</p>
                 <h2 className="text-3xl font-black text-slate-900">GH₵ {stationWallet.escrow.toLocaleString('en-US', { minimumFractionDigits: 2 })}</h2>
                 <p className="text-[11px] text-slate-500 mt-2 leading-relaxed max-w-[85%]">
                   Funds generated from 95% split ticket sales. Requires bus departure confirmation to release into Available Payouts. Operator cancellations trigger automated 90% customer refunds.
                 </p>
               </div>
            </div>
          </div>
        </div>
      )}

      {activeTab === 'dispatch' && (
        <div className="p-6 max-w-md mx-auto">
           <h3 className="font-bold text-slate-900 text-sm uppercase tracking-wider mb-4">Pending Departures & Escrow Locks</h3>
           
           <div className="space-y-4">
             {trips.filter(t => ['Scheduled', 'Boarding'].includes(t.status)).length === 0 && (
                <div className="text-center py-10 bg-white rounded-3xl border border-slate-100">
                  <CheckCircle size={32} className="mx-auto text-emerald-300 mb-2" />
                  <p className="text-sm font-bold text-slate-600">All buses departed.</p>
                </div>
             )}

             {trips.filter(t => ['Scheduled', 'Boarding'].includes(t.status)).map(trip => {
                return (
                  <div key={trip.id} className="bg-white p-5 rounded-[24px] shadow-sm border border-slate-200 flex flex-col gap-4">
                    <div className="flex justify-between items-start">
                        <div>
                          <span className={`text-[10px] px-2 py-0.5 rounded-full font-bold uppercase tracking-wider mb-2 inline-block ${trip.status === 'Boarding' ? 'bg-amber-100 text-amber-700' : 'bg-slate-100 text-slate-600'}`}>{trip.status}</span>
                          <p className="font-black text-slate-900 text-lg">{trip.departureTime}</p>
                          <p className="text-xs text-slate-500 font-medium">To {trip.to} • {trip.bookedSeats.length} Pax</p>
                        </div>
                        <div className="text-right bg-amber-50 px-3 py-2 rounded-xl border border-amber-100">
                          <p className="text-[10px] text-amber-700 uppercase font-bold flex items-center gap-1 justify-end"><Lock size={12}/> Escrow Held</p>
                          <p className="font-black text-amber-600 text-base mt-0.5">GH₵ {(trip.escrowHeld || 0).toFixed(2)}</p>
                        </div>
                    </div>
                    
                    <div className="flex gap-2">
                      <Button
                          fullWidth
                          variant="primary"
                          icon={Unlock}
                          onClick={() => handleDispatchBus(trip.id)}
                      >
                          Confirm Departure
                      </Button>
                      <Button
                          variant="danger"
                          icon={ShieldAlert}
                          onClick={() => handleCancelTripByOperator(trip.id)}
                          title="Cancel Trip & Refund 90%"
                        >
                          Cancel & Refund
                      </Button>
                    </div>
                  </div>
                );
             })}
           </div>
        </div>
      )}

      {showSignModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-[32px] p-6 max-w-md w-full space-y-4 shadow-2xl animate-fadeIn">
            <div className="flex justify-between items-center border-b border-slate-100 pb-3">
              <h3 className="font-black text-slate-900 flex items-center gap-2"><FileText size={18} className="text-indigo-600" /> Digital Operator Terms Agreement</h3>
              <button onClick={() => setShowSignModal(false)} className="text-slate-400 hover:text-slate-600 font-bold">✕</button>
            </div>
            
            <div className="bg-slate-50 p-4 rounded-2xl border border-slate-200 text-xs text-slate-600 space-y-2 max-h-48 overflow-y-auto leading-relaxed">
              <p className="font-bold text-slate-900">LIABILITY WAIVER & CHARGEBACK SHIELD PROTOCOL</p>
              <p>1. The transport operator assumes full legal and financial responsibility for any service cancellation, mechanical breakdown, passenger injury, or luggage loss.</p>
              <p>2. The software platform (EasyFare OS) acts solely as a technology intermediary via Paystack Split Payments. Any passenger chargeback or dispute will be directly offset against operator escrow holds.</p>
              <p>3. By signing below, the operator binds their corporate entity and assigns full liability away from the platform creator.</p>
            </div>

            <Input label="Authorized Representative Name" placeholder="e.g. Kofi Annan" value={operatorRepName} onChange={e => setOperatorRepName(e.target.value)} />
            <Input label="National ID / Passport Number" placeholder="e.g. GHA-892104921-1" value={operatorIdNumber} onChange={e => setOperatorIdNumber(e.target.value)} />

            <div className="flex gap-2 pt-2">
              <Button variant="outline" fullWidth onClick={() => setShowSignModal(false)}>Cancel</Button>
              <Button fullWidth icon={CheckCircle} onClick={handleSignTerms}>Sign & Accept Terms</Button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default function App() {
  const [userRole, setUserRole] = useState(() => localStorage.getItem(STORAGE_KEY_AUTH) || null);
  const [currentUser, setCurrentUser] = useState(() => {
    try { const saved = localStorage.getItem(STORAGE_KEY_USER); return saved ? JSON.parse(saved) : null; } catch { return null; }
  });

  const [currentView, setCurrentView] = useState('home');
  const [searchParams, setSearchParams] = useState(null);
  const [selectedTrip, setSelectedTrip] = useState(null);
  const [selectedSeats, setSelectedSeats] = useState([]);
  const [parcelDetails, setParcelDetails] = useState(null);
  
  const [stationWallet, setStationWallet] = useState(() => {
    try { const saved = localStorage.getItem(STORAGE_KEY_WALLET); return saved ? JSON.parse(saved) : INITIAL_WALLET; } catch { return INITIAL_WALLET; }
  });

  const [platformWallet, setPlatformWallet] = useState(() => {
    try { const saved = localStorage.getItem(STORAGE_KEY_PLATFORM_WALLET); return saved ? JSON.parse(saved) : INITIAL_PLATFORM_WALLET; } catch { return INITIAL_PLATFORM_WALLET; }
  });

  const [trips, setTrips] = useState(() => {
    try { const saved = localStorage.getItem(STORAGE_KEY_TRIPS); return saved ? JSON.parse(saved) : initialTrips; } catch { return initialTrips; }
  });

  const [myTickets, setMyTickets] = useState(() => {
    try { const saved = localStorage.getItem(STORAGE_KEY_TICKETS); return saved ? JSON.parse(saved) : []; } catch { return []; }
  });

  const [myParcels, setMyParcels] = useState(() => {
    try { const saved = localStorage.getItem(STORAGE_KEY_PARCELS); return saved ? JSON.parse(saved) : []; } catch { return []; }
  });

  const [refunds, setRefunds] = useState(() => {
    try { const saved = localStorage.getItem(STORAGE_KEY_REFUNDS); return saved ? JSON.parse(saved) : []; } catch { return []; }
  });

  const [disputes, setDisputes] = useState(() => {
    try { const saved = localStorage.getItem(STORAGE_KEY_DISPUTES); return saved ? JSON.parse(saved) : []; } catch { return []; }
  });

  const [manifestLedger, setManifestLedger] = useState(() => {
    try { const saved = localStorage.getItem(STORAGE_KEY_MANIFEST); return saved ? JSON.parse(saved) : []; } catch { return []; }
  });

  const [toast, setToast] = useState(null);

  useEffect(() => { try { localStorage.setItem(STORAGE_KEY_WALLET, JSON.stringify(stationWallet)); } catch (e) {} }, [stationWallet]);
  useEffect(() => { try { localStorage.setItem(STORAGE_KEY_PLATFORM_WALLET, JSON.stringify(platformWallet)); } catch (e) {} }, [platformWallet]);
  useEffect(() => { try { localStorage.setItem(STORAGE_KEY_TICKETS, JSON.stringify(myTickets)); } catch (e) {} }, [myTickets]);
  useEffect(() => { try { localStorage.setItem(STORAGE_KEY_PARCELS, JSON.stringify(myParcels)); } catch (e) {} }, [myParcels]);
  useEffect(() => { try { localStorage.setItem(STORAGE_KEY_REFUNDS, JSON.stringify(refunds)); } catch (e) {} }, [refunds]);
  useEffect(() => { try { localStorage.setItem(STORAGE_KEY_DISPUTES, JSON.stringify(disputes)); } catch (e) {} }, [disputes]);
  useEffect(() => { try { localStorage.setItem(STORAGE_KEY_MANIFEST, JSON.stringify(manifestLedger)); } catch (e) {} }, [manifestLedger]);
  useEffect(() => { try { localStorage.setItem(STORAGE_KEY_TRIPS, JSON.stringify(trips)); } catch (e) {} }, [trips]);
  useEffect(() => { userRole ? localStorage.setItem(STORAGE_KEY_AUTH, userRole) : localStorage.removeItem(STORAGE_KEY_AUTH); }, [userRole]);
  useEffect(() => { currentUser ? localStorage.setItem(STORAGE_KEY_USER, JSON.stringify(currentUser)) : localStorage.removeItem(STORAGE_KEY_USER); }, [currentUser]);

  if (!userRole) {
    return (
      <div className="w-full max-w-md mx-auto min-h-screen bg-slate-50 relative shadow-2xl overflow-hidden border-x border-slate-200">
        {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}
        <AuthScreen 
          setToast={setToast} 
          onAuthSuccess={(role, userObj) => { 
            setUserRole(role); setCurrentUser(userObj); setCurrentView('home'); 
            setToast({ message: `Welcome back!`, type: 'success' }); 
          }} 
        />
      </div>
    );
  }

  const handleCheckoutComplete = (checkoutData) => {
    const platformFee = checkoutData.financials.platformFee;
    const netPayout = checkoutData.financials.netCompanyPayout;
    
    setPlatformWallet(prev => ({
      balance: prev.balance + platformFee,
      totalSplitCollected: prev.totalSplitCollected + platformFee
    }));

    setStationWallet(prev => ({ ...prev, escrow: prev.escrow + netPayout }));

    const updatedTrips = trips.map(t => {
      if (t.id === selectedTrip.id) {
        return {
          ...t,
          bookedSeats: [...t.bookedSeats, ...selectedSeats],
          escrowHeld: (t.escrowHeld || 0) + netPayout
        };
      }
      return t;
    });
    setTrips(updatedTrips);
    
    const updatedSelectedTrip = updatedTrips.find(t => t.id === selectedTrip.id);
    setSelectedTrip(updatedSelectedTrip);

    const newTicket = { 
      trip: updatedSelectedTrip, seats: selectedSeats, luggage: checkoutData.luggage,
      financials: checkoutData.financials, bookedAt: new Date().toISOString(), boardingVerified: false 
    };
    
    setMyTickets([newTicket, ...myTickets]);
    setToast({ message: 'Paystack Split Executed! 5% Creator fee routed & 95% Escrow locked.', type: 'success' });
    setCurrentView('tickets');
    setSelectedSeats([]);
  };

  const handleCancelTicket = (ticketIdx) => {
    const ticketToCancel = myTickets[ticketIdx];
    const refundAmount = ticketToCancel.financials?.netCompanyPayout || 0;
    
    const updatedTrips = trips.map(t => {
      if (t.id === ticketToCancel.trip.id) {
        const isDeparted = t.status === 'En Route';
        return {
          ...t,
          bookedSeats: t.bookedSeats.filter(s => !ticketToCancel.seats.includes(s)),
          escrowHeld: isDeparted ? t.escrowHeld : Math.max(0, (t.escrowHeld || 0) - refundAmount)
        };
      }
      return t;
    });
    setTrips(updatedTrips);

    if (ticketToCancel.trip.status !== 'En Route') {
       setStationWallet(prev => ({ ...prev, escrow: Math.max(0, prev.escrow - refundAmount) }));
    }

    const updatedTickets = myTickets.filter((_, idx) => idx !== ticketIdx);
    setMyTickets(updatedTickets);
    setToast({ message: 'Ticket cancelled. Escrow hold refunded.', type: 'success' });
  };

  const handleRaiseDispute = (disputeObj) => {
    setDisputes(prev => [disputeObj, ...prev]);
  };

  const handleParcelComplete = (parcelObj) => {
    const platformFee = parseFloat((parcelObj.price * 0.05).toFixed(2));
    const stationAmount = parcelObj.price - platformFee;

    setPlatformWallet(prev => ({
      balance: prev.balance + platformFee,
      totalSplitCollected: prev.totalSplitCollected + platformFee
    }));

    setStationWallet(prev => ({ ...prev, escrow: prev.escrow + stationAmount }));

    setMyParcels([parcelObj, ...myParcels]);
    setToast({ message: 'Parcel dispatched. Paystack Split 5% credited instantly!', type: 'success' });
    setCurrentView('tickets');
  };

  const CustomerTabBar = () => (
    <div className="fixed bottom-0 left-0 right-0 bg-white border-t border-slate-200 flex justify-around items-center py-2 px-2 shadow-[0_-5px_20px_rgba(0,0,0,0.05)] z-40 max-w-md mx-auto">
      {[
        { id: 'home', label: 'Explore', icon: Search },
        { id: 'tickets', label: 'Passes', icon: Ticket, badge: myTickets.length + myParcels.length + disputes.length },
        { id: 'profile', label: 'Profile', icon: User }
      ].map(tab => {
        const Icon = tab.icon;
        const isActive = currentView === tab.id;
        return (
          <button 
            key={tab.id} onClick={() => setCurrentView(tab.id)} 
            className={`flex flex-col items-center p-2 rounded-2xl w-16 transition-all ${isActive ? 'text-indigo-600 bg-indigo-50/50' : 'text-slate-400 hover:text-slate-600'}`}
          >
            <div className="relative">
              <Icon size={22} className={isActive ? 'fill-indigo-100' : ''} />
              {tab.badge > 0 && <div className="absolute -top-1 -right-1 w-2.5 h-2.5 bg-rose-500 rounded-full border-2 border-white"></div>}
            </div>
            <span className="text-[10px] font-bold mt-1 tracking-wider uppercase">{tab.label}</span>
          </button>
        );
      })}
    </div>
  );

  return (
    <div className="w-full max-w-md mx-auto min-h-screen bg-slate-50 relative shadow-2xl overflow-hidden border-x border-slate-200">
      {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}
      
      {userRole === 'admin' && (
        <AdminDashboard 
          stationWallet={stationWallet} setStationWallet={setStationWallet}
          platformWallet={platformWallet} setPlatformWallet={setPlatformWallet}
          trips={trips} setTrips={setTrips}
          refunds={refunds} setRefunds={setRefunds}
          disputes={disputes} setDisputes={setDisputes}
          myTickets={myTickets} setMyTickets={setMyTickets}
          manifestLedger={manifestLedger} setManifestLedger={setManifestLedger}
          onLogout={() => { setUserRole(null); setCurrentUser(null); }} 
          setToast={setToast}
        />
      )}

      {userRole === 'customer' && (
        <>
          {currentView === 'home' && <CustomerSearch onSearch={(params) => { setSearchParams(params); setCurrentView('searchResults'); }} onSelectParcel={(params) => { setParcelDetails(params); setCurrentView('parcelBooking'); }} setToast={setToast} />}
          {currentView === 'searchResults' && <SearchResults searchParams={searchParams} trips={trips} onBack={() => setCurrentView('home')} onSelectTrip={(trip) => { setSelectedTrip(trip); setCurrentView('seatSelect'); }} />}
          {currentView === 'seatSelect' && <SeatSelection trip={selectedTrip} onBack={() => setCurrentView('searchResults')} onContinue={(seats) => { setSelectedSeats(seats); setCurrentView('checkout'); }} setToast={setToast} />}
          {currentView === 'checkout' && <Checkout trip={selectedTrip} seats={selectedSeats} onBack={() => setCurrentView('seatSelect')} onComplete={handleCheckoutComplete} setToast={setToast} />}
          {currentView === 'parcelBooking' && <ParcelBooking details={parcelDetails} onBack={() => setCurrentView('home')} onComplete={handleParcelComplete} setToast={setToast} />}
          {currentView === 'liveTracker' && <LiveBusTracker trip={selectedTrip || trips[0]} onBack={() => setCurrentView('tickets')} setToast={setToast} />}
          {currentView === 'tickets' && <MyTickets tickets={myTickets} parcels={myParcels} refunds={refunds} disputes={disputes} onTrackBus={(trip) => { setSelectedTrip(trip); setCurrentView('liveTracker'); }} onCancelTicket={handleCancelTicket} onRaiseDispute={handleRaiseDispute} setToast={setToast} />}
          {currentView === 'profile' && (
             <div className="p-6 max-w-md mx-auto pb-32">
                <h1 className="text-2xl font-black text-slate-900 mb-6">Passenger Profile</h1>
                <div className="bg-white rounded-[32px] p-6 shadow-sm border border-slate-100 mb-6 flex items-center gap-4">
                  <div className="w-16 h-16 bg-indigo-100 text-indigo-700 rounded-2xl flex items-center justify-center font-black text-xl">
                    {currentUser?.name ? currentUser.name.substring(0, 2).toUpperCase() : 'JM'}
                  </div>
                  <div>
                    <h2 className="font-bold text-slate-900 text-base">{currentUser?.name || 'John Mensah'}</h2>
                    <p className="text-slate-500 text-xs">{currentUser?.phone || currentUser?.email || '+233 54 123 4567'} • Verified</p>
                  </div>
                </div>
                <Button variant="danger" fullWidth icon={LogOut} onClick={() => { setUserRole(null); setCurrentUser(null); }}>
                  Sign Out Securely
                </Button>
             </div>
          )}

          {['home', 'tickets', 'profile'].includes(currentView) && <CustomerTabBar />}
        </>
      )}
    </div>
  );
}