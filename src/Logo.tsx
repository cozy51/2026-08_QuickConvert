import { useId } from 'react';

// ブランドマーク。public/favicon.svg と同一の形状（スクワークル + 稲妻）。
export default function Logo({ className }: { className?: string }) {
 const uid=useId(); const badge=`qc-badge-${uid}`; const spot=`qc-spot-${uid}`;
 const squircle='M19 0H45C57.92 0 64 6.08 64 19V45C64 57.92 57.92 64 45 64H19C6.08 64 0 57.92 0 45V19C0 6.08 6.08 0 19 0Z';
 return <svg className={className} viewBox="0 0 64 64" role="img" aria-label="QuickConvert">
  <defs>
   <linearGradient id={badge} x1="4" y1="0" x2="60" y2="64" gradientUnits="userSpaceOnUse"><stop stopColor="#6E63F0"/><stop offset="1" stopColor="#9A55E2"/></linearGradient>
   <radialGradient id={spot} cx="0" cy="0" r="1" gradientTransform="translate(15 7) rotate(48) scale(56)"><stop stopColor="#fff" stopOpacity=".30"/><stop offset="1" stopColor="#fff" stopOpacity="0"/></radialGradient>
  </defs>
  <path d={squircle} fill={`url(#${badge})`}/>
  <path d={squircle} fill={`url(#${spot})`}/>
  <path d={squircle} fill="none" stroke="#fff" strokeOpacity=".22" strokeWidth="1.22" transform="translate(.6 .6) scale(.98125)"/>
  <path d="M38.82 8.39 L18.36 35.37 Q17.5 36.5 18.92 36.5 L27 36.5 Q29 36.5 28.56 38.45 L24.75 55.41 Q24.5 56.5 25.18 55.61 L45.64 28.63 Q46.5 27.5 45.08 27.5 L37 27.5 Q35 27.5 35.44 25.55 L39.25 8.59 Q39.5 7.5 38.82 8.39 Z" fill="#fff"/>
 </svg>;
}
