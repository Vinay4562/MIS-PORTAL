import { createContext, useContext, useState, useEffect } from 'react';
import axios from 'axios';

const AuthContext = createContext({});

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

export function AuthProvider({ children }) {
  const [user, setUser] = useState(null);
  const [loading, setLoading] = useState(true);
  const [token, setToken] = useState(localStorage.getItem('token'));

  const logout = () => {
    localStorage.removeItem('token');
    setToken(null);
    setUser(null);
    delete axios.defaults.headers.common['Authorization'];
  };

  useEffect(() => {
    const initAuth = async () => {
      if (!token) {
        setLoading(false);
        return;
      }
      axios.defaults.headers.common['Authorization'] = `Bearer ${token}`;
      try {
        const response = await axios.get(`${API}/auth/me`);
        setUser(response.data);
      } catch (error) {
        console.error('Failed to fetch user:', error);
        logout();
      } finally {
        setLoading(false);
      }
    };

    initAuth();
  }, [token]);

  const login = async (email, password) => {
    const response = await axios.post(`${API}/auth/login`, { email, password }, { timeout: 15000 });
    const { access_token, user } = response.data;
    localStorage.setItem('token', access_token);
    setToken(access_token);
    setUser(user);
    axios.defaults.headers.common['Authorization'] = `Bearer ${access_token}`;
    return user;
  };

  const register = async (email, password, full_name) => {
    const response = await axios.post(`${API}/auth/register`, { email, password, full_name }, { timeout: 15000 });
    const { access_token, user } = response.data;
    localStorage.setItem('token', access_token);
    setToken(access_token);
    setUser(user);
    axios.defaults.headers.common['Authorization'] = `Bearer ${access_token}`;
    return user;
  };

  const signupRequest = async (email, password, full_name) => {
    await axios.post(`${API}/auth/signup-request`, { email, password, full_name }, { timeout: 15000 });
  };

  const signupVerify = async (email, otp) => {
    const response = await axios.post(`${API}/auth/signup-verify`, { email, otp });
    const { access_token, user } = response.data;
    localStorage.setItem('token', access_token);
    setToken(access_token);
    setUser(user);
    axios.defaults.headers.common['Authorization'] = `Bearer ${access_token}`;
    return user;
  };

  return (
    <AuthContext.Provider value={{ user, loading, login, register, logout, signupRequest, signupVerify }}>
      {children}
    </AuthContext.Provider>
  );
}

export const useAuth = () => useContext(AuthContext);
