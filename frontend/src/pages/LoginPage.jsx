import { useState } from 'react';
import { useNavigate, Link } from 'react-router-dom';
import { useAuth } from '@/contexts/AuthContext';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { toast } from 'sonner';
import { Zap, Eye, EyeOff } from 'lucide-react';

export default function LoginPage() {
  const [isLogin, setIsLogin] = useState(true);
  const [isSignupVerification, setIsSignupVerification] = useState(false);
  const [showPassword, setShowPassword] = useState(false);
  const [otp, setOtp] = useState('');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [fullName, setFullName] = useState('');
  const [loading, setLoading] = useState(false);
  const { login, signupRequest, signupVerify } = useAuth();
  const navigate = useNavigate();

  const handleSubmit = async (e) => {
    e.preventDefault();
    setLoading(true);
    
    try {
      if (isLogin) {
        await login(email, password);
        toast.success('Login successful!');
        navigate('/');
      } else {
        if (!isSignupVerification) {
          if (!fullName.trim()) {
            toast.error('Full name is required');
            setLoading(false);
            return;
          }
          await signupRequest(email, password, fullName);
          setIsSignupVerification(true);
          toast.success('OTP sent to admin. Please enter the OTP to verify.');
          setLoading(false);
          return;
        } else {
          if (!otp.trim()) {
            toast.error('OTP is required');
            setLoading(false);
            return;
          }
          await signupVerify(email, otp);
          toast.success('Registration successful!');
          navigate('/');
        }
      }
    } catch (error) {
      toast.error(error.response?.data?.detail || 'Authentication failed');
      setLoading(false);
    } finally {
      if (isLogin || (isSignupVerification && otp)) { 
        // Only stop loading if we are navigating or failed. 
        // If we just switched to verification mode, we already stopped loading above.
        // But to be safe, we can just set loading false here if not navigated.
        // Actually, navigate happens above. If we are here, it means error or stopped.
        // But if we navigated, this component unmounts.
        // If we switched mode, we set loading false manually.
        // So this finally block is fine to ensure loading is off on error.
      }
      if (!isSignupVerification || otp) {
         setLoading(false);
      }
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-slate-50 dark:bg-slate-900 p-4"
         style={{
           backgroundImage: 'url(https://images.unsplash.com/photo-1757866419834-192728bdb138?crop=entropy&cs=srgb&fm=jpg&ixid=M3w4NjA3MDB8MHwxfHNlYXJjaHwyfHxhYnN0cmFjdCUyMGVsZWN0cmljJTIwcG93ZXIlMjBsaW5lcyUyMG1pbmltYWxpc3R8ZW58MHx8fHwxNzcwMDMzNTU4fDA&ixlib=rb-4.1.0&q=85)',
           backgroundSize: 'cover',
           backgroundPosition: 'center'
         }}>
      <div className="absolute inset-0 bg-white/90 dark:bg-slate-900/90 backdrop-blur-sm"></div>
      
      <Card className="w-full max-w-md relative z-10 shadow-xl border-2 bg-white dark:bg-slate-800">
        <CardHeader className="space-y-3 bg-white dark:bg-slate-800">
          <div className="flex items-center justify-center mb-2">
            <div className="p-3 bg-blue-600 rounded-lg">
              <Zap className="w-8 h-8 text-white" />
            </div>
          </div>
          <CardTitle className="text-3xl font-heading text-center text-slate-900 dark:text-slate-100">
            MIS PORTAL
          </CardTitle>
          <CardDescription className="text-center text-base text-slate-600 dark:text-slate-400">
            {isLogin ? 'Sign in to your account' : 'Create a new account'}
          </CardDescription>
        </CardHeader>
        <CardContent>
          <form onSubmit={handleSubmit} className="space-y-4">
            {!isLogin && !isSignupVerification && (
              <div className="space-y-2">
                <Label htmlFor="fullName" data-testid="full-name-label">Full Name</Label>
                <Input
                  id="fullName"
                  data-testid="full-name-input"
                  type="text"
                  placeholder="John Doe"
                  value={fullName}
                  onChange={(e) => setFullName(e.target.value)}
                  required={!isLogin}
                />
              </div>
            )}
            
            {isSignupVerification ? (
              <div className="space-y-2">
                <div className="p-4 bg-blue-50 dark:bg-blue-900/20 rounded-lg mb-4 text-sm text-blue-700 dark:text-blue-300">
                  An OTP has been sent to the admin for approval. Please enter the OTP below to complete your registration.
                </div>
                <Label htmlFor="otp">OTP Code</Label>
                <Input 
                  id="otp" 
                  value={otp} 
                  onChange={(e) => setOtp(e.target.value)} 
                  placeholder="Enter 6-digit OTP"
                  required 
                />
              </div>
            ) : (
              <>
                <div className="space-y-2">
                  <Label htmlFor="email" data-testid="email-label">Email</Label>
                  <Input
                    id="email"
                    data-testid="email-input"
                    type="email"
                    placeholder="you@example.com"
                    value={email}
                    onChange={(e) => setEmail(e.target.value)}
                    required
                  />
                </div>
                
                <div className="space-y-2">
                  <Label htmlFor="password" data-testid="password-label">Password</Label>
                  <div className="relative">
                    <Input
                      id="password"
                      data-testid="password-input"
                      type={showPassword ? "text" : "password"}
                      placeholder="••••••••"
                      value={password}
                      onChange={(e) => setPassword(e.target.value)}
                      required
                      className="pr-10"
                    />
                    <button
                      type="button"
                      onClick={() => setShowPassword(!showPassword)}
                      className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-200"
                    >
                      {showPassword ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
                    </button>
                  </div>
                  {isLogin && (
                    <div className="text-right">
                      <Link 
                        to="/forgot-password" 
                        className="text-sm text-blue-600 hover:text-blue-700 dark:text-blue-400 dark:hover:text-blue-300"
                      >
                        Forgot Password?
                      </Link>
                    </div>
                  )}
                </div>
              </>
            )}
            
            <Button 
              type="submit" 
              className="w-full" 
              disabled={loading}
              data-testid="submit-button"
            >
              {loading ? 'Please wait...' : (isLogin ? 'Sign In' : (isSignupVerification ? 'Verify & Create Account' : 'Sign Up'))}
            </Button>
            
            <div className="text-center pt-2">
              <button
                type="button"
                onClick={() => {
                  setIsLogin(!isLogin);
                  setIsSignupVerification(false);
                  setOtp('');
                }}
                className="text-sm text-blue-600 hover:text-blue-700 dark:text-blue-400 dark:hover:text-blue-300 font-medium"
                data-testid="toggle-auth-mode"
              >
                {isLogin ? "Don't have an account? Sign up" : 'Already have an account? Sign in'}
              </button>
            </div>
          </form>
        </CardContent>
      </Card>
    </div>
  );
}
