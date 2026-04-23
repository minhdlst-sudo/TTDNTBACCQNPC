import React from 'react';

/**
 * Logo đặc trưng cho ứng dụng Giảm Tổn thất điện năng.
 * Kết hợp biểu tượng tia sét (năng lượng) và biểu đồ xu thế (giảm thiểu).
 */
export const AppLogo = ({ className = "w-10 h-10" }: { className?: string }) => {
  return (
    <svg 
      viewBox="0 0 100 100" 
      fill="none" 
      xmlns="http://www.w3.org/2000/svg" 
      className={className}
    >
      {/* Nền tròn hiện đại */}
      <circle cx="50" cy="50" r="48" fill="url(#bgGradient)" />
      
      {/* Vòng tròn giám sát */}
      <circle cx="50" cy="50" r="38" stroke="white" strokeWidth="2" strokeDasharray="10 5" opacity="0.3" />
      
      {/* Biểu tượng tia sét chính */}
      <path 
        d="M55 20L35 50H45L40 80L65 40H52L55 20Z" 
        fill="white" 
        filter="url(#glow)"
      />
      
      {/* Đường biểu đồ giảm tổn thất */}
      <path 
        d="M30 40Q40 45 50 40T75 30" 
        stroke="#10b981" 
        strokeWidth="4" 
        strokeLinecap="round" 
        opacity="0.8" 
      />
      <path 
        d="M70 25L75 30L70 35" 
        stroke="#10b981" 
        strokeWidth="4" 
        strokeLinecap="round" 
        strokeLinejoin="round" 
      />

      <defs>
        <linearGradient id="bgGradient" x1="0" y1="0" x2="100" y2="100" gradientUnits="userSpaceOnUse">
          <stop stopColor="#1e40af" />
          <stop offset="1" stopColor="#3b82f6" />
        </linearGradient>
        <filter id="glow" x="-20%" y="-20%" width="140%" height="140%">
          <stop offset="0" stopColor="white" stopOpacity="0.8" />
          <feGaussianBlur stdDeviation="2" result="blur" />
          <feComposite in="SourceGraphic" in2="blur" operator="over" />
        </filter>
      </defs>
    </svg>
  );
};

export default AppLogo;
