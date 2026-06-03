import * as React from 'react';
import { runCrmImport } from './crmImport';
import type { CrmPhone, CrmEmail, CrmPerson, CrmCompany, CrmActivity, CrmActivityType } from './crmTypes';

export type { CrmPerson, CrmCompany, CrmActivity } from './crmTypes';

// ── Constants ─────────────────────────────────────────────────────
const PHONE_TYPES  = ['Work', 'Home', 'Mobile', 'Other'];
const EMAIL_TYPES  = ['Work', 'Home', 'Other'];

const COUNTRY_CODES = [
  { cc: '+93',   flag: '🇦🇫', name: 'Afghanistan'                        },
  { cc: '+355',  flag: '🇦🇱', name: 'Albania'                            },
  { cc: '+213',  flag: '🇩🇿', name: 'Algeria'                            },
  { cc: '+376',  flag: '🇦🇩', name: 'Andorra'                            },
  { cc: '+244',  flag: '🇦🇴', name: 'Angola'                             },
  { cc: '+1268', flag: '🇦🇬', name: 'Antigua and Barbuda'                },
  { cc: '+54',   flag: '🇦🇷', name: 'Argentina'                          },
  { cc: '+374',  flag: '🇦🇲', name: 'Armenia'                            },
  { cc: '+61',   flag: '🇦🇺', name: 'Australia'                          },
  { cc: '+43',   flag: '🇦🇹', name: 'Austria'                            },
  { cc: '+994',  flag: '🇦🇿', name: 'Azerbaijan'                         },
  { cc: '+1242', flag: '🇧🇸', name: 'Bahamas'                            },
  { cc: '+973',  flag: '🇧🇭', name: 'Bahrain'                            },
  { cc: '+880',  flag: '🇧🇩', name: 'Bangladesh'                         },
  { cc: '+1246', flag: '🇧🇧', name: 'Barbados'                           },
  { cc: '+375',  flag: '🇧🇾', name: 'Belarus'                            },
  { cc: '+32',   flag: '🇧🇪', name: 'Belgium'                            },
  { cc: '+501',  flag: '🇧🇿', name: 'Belize'                             },
  { cc: '+229',  flag: '🇧🇯', name: 'Benin'                              },
  { cc: '+975',  flag: '🇧🇹', name: 'Bhutan'                             },
  { cc: '+591',  flag: '🇧🇴', name: 'Bolivia'                            },
  { cc: '+387',  flag: '🇧🇦', name: 'Bosnia and Herzegovina'             },
  { cc: '+267',  flag: '🇧🇼', name: 'Botswana'                           },
  { cc: '+55',   flag: '🇧🇷', name: 'Brazil'                             },
  { cc: '+673',  flag: '🇧🇳', name: 'Brunei'                             },
  { cc: '+359',  flag: '🇧🇬', name: 'Bulgaria'                           },
  { cc: '+226',  flag: '🇧🇫', name: 'Burkina Faso'                       },
  { cc: '+257',  flag: '🇧🇮', name: 'Burundi'                            },
  { cc: '+238',  flag: '🇨🇻', name: 'Cabo Verde'                         },
  { cc: '+855',  flag: '🇰🇭', name: 'Cambodia'                           },
  { cc: '+237',  flag: '🇨🇲', name: 'Cameroon'                           },
  { cc: '+1',    flag: '🇨🇦', name: 'Canada'                             },
  { cc: '+236',  flag: '🇨🇫', name: 'Central African Republic'           },
  { cc: '+235',  flag: '🇹🇩', name: 'Chad'                               },
  { cc: '+56',   flag: '🇨🇱', name: 'Chile'                              },
  { cc: '+86',   flag: '🇨🇳', name: 'China'                              },
  { cc: '+57',   flag: '🇨🇴', name: 'Colombia'                           },
  { cc: '+269',  flag: '🇰🇲', name: 'Comoros'                            },
  { cc: '+243',  flag: '🇨🇩', name: 'Congo (DRC)'                        },
  { cc: '+242',  flag: '🇨🇬', name: 'Congo (Republic)'                   },
  { cc: '+506',  flag: '🇨🇷', name: 'Costa Rica'                         },
  { cc: '+385',  flag: '🇭🇷', name: 'Croatia'                            },
  { cc: '+53',   flag: '🇨🇺', name: 'Cuba'                               },
  { cc: '+357',  flag: '🇨🇾', name: 'Cyprus'                             },
  { cc: '+420',  flag: '🇨🇿', name: 'Czech Republic'                     },
  { cc: '+45',   flag: '🇩🇰', name: 'Denmark'                            },
  { cc: '+253',  flag: '🇩🇯', name: 'Djibouti'                           },
  { cc: '+1767', flag: '🇩🇲', name: 'Dominica'                           },
  { cc: '+1809', flag: '🇩🇴', name: 'Dominican Republic'                 },
  { cc: '+593',  flag: '🇪🇨', name: 'Ecuador'                            },
  { cc: '+20',   flag: '🇪🇬', name: 'Egypt'                              },
  { cc: '+503',  flag: '🇸🇻', name: 'El Salvador'                        },
  { cc: '+240',  flag: '🇬🇶', name: 'Equatorial Guinea'                  },
  { cc: '+291',  flag: '🇪🇷', name: 'Eritrea'                            },
  { cc: '+372',  flag: '🇪🇪', name: 'Estonia'                            },
  { cc: '+268',  flag: '🇸🇿', name: 'Eswatini'                           },
  { cc: '+251',  flag: '🇪🇹', name: 'Ethiopia'                           },
  { cc: '+679',  flag: '🇫🇯', name: 'Fiji'                               },
  { cc: '+358',  flag: '🇫🇮', name: 'Finland'                            },
  { cc: '+33',   flag: '🇫🇷', name: 'France'                             },
  { cc: '+241',  flag: '🇬🇦', name: 'Gabon'                              },
  { cc: '+220',  flag: '🇬🇲', name: 'Gambia'                             },
  { cc: '+995',  flag: '🇬🇪', name: 'Georgia'                            },
  { cc: '+49',   flag: '🇩🇪', name: 'Germany'                            },
  { cc: '+233',  flag: '🇬🇭', name: 'Ghana'                              },
  { cc: '+30',   flag: '🇬🇷', name: 'Greece'                             },
  { cc: '+1473', flag: '🇬🇩', name: 'Grenada'                            },
  { cc: '+502',  flag: '🇬🇹', name: 'Guatemala'                          },
  { cc: '+224',  flag: '🇬🇳', name: 'Guinea'                             },
  { cc: '+245',  flag: '🇬🇼', name: 'Guinea-Bissau'                      },
  { cc: '+592',  flag: '🇬🇾', name: 'Guyana'                             },
  { cc: '+509',  flag: '🇭🇹', name: 'Haiti'                              },
  { cc: '+504',  flag: '🇭🇳', name: 'Honduras'                           },
  { cc: '+852',  flag: '🇭🇰', name: 'Hong Kong'                          },
  { cc: '+36',   flag: '🇭🇺', name: 'Hungary'                            },
  { cc: '+354',  flag: '🇮🇸', name: 'Iceland'                            },
  { cc: '+91',   flag: '🇮🇳', name: 'India'                              },
  { cc: '+62',   flag: '🇮🇩', name: 'Indonesia'                          },
  { cc: '+98',   flag: '🇮🇷', name: 'Iran'                               },
  { cc: '+964',  flag: '🇮🇶', name: 'Iraq'                               },
  { cc: '+353',  flag: '🇮🇪', name: 'Ireland'                            },
  { cc: '+972',  flag: '🇮🇱', name: 'Israel'                             },
  { cc: '+39',   flag: '🇮🇹', name: 'Italy'                              },
  { cc: '+1876', flag: '🇯🇲', name: 'Jamaica'                            },
  { cc: '+81',   flag: '🇯🇵', name: 'Japan'                              },
  { cc: '+962',  flag: '🇯🇴', name: 'Jordan'                             },
  { cc: '+7',    flag: '🇰🇿', name: 'Kazakhstan'                         },
  { cc: '+254',  flag: '🇰🇪', name: 'Kenya'                              },
  { cc: '+686',  flag: '🇰🇮', name: 'Kiribati'                           },
  { cc: '+383',  flag: '🇽🇰', name: 'Kosovo'                             },
  { cc: '+965',  flag: '🇰🇼', name: 'Kuwait'                             },
  { cc: '+996',  flag: '🇰🇬', name: 'Kyrgyzstan'                         },
  { cc: '+856',  flag: '🇱🇦', name: 'Laos'                               },
  { cc: '+371',  flag: '🇱🇻', name: 'Latvia'                             },
  { cc: '+961',  flag: '🇱🇧', name: 'Lebanon'                            },
  { cc: '+266',  flag: '🇱🇸', name: 'Lesotho'                            },
  { cc: '+231',  flag: '🇱🇷', name: 'Liberia'                            },
  { cc: '+218',  flag: '🇱🇾', name: 'Libya'                              },
  { cc: '+423',  flag: '🇱🇮', name: 'Liechtenstein'                      },
  { cc: '+370',  flag: '🇱🇹', name: 'Lithuania'                          },
  { cc: '+352',  flag: '🇱🇺', name: 'Luxembourg'                         },
  { cc: '+853',  flag: '🇲🇴', name: 'Macau'                              },
  { cc: '+261',  flag: '🇲🇬', name: 'Madagascar'                         },
  { cc: '+265',  flag: '🇲🇼', name: 'Malawi'                             },
  { cc: '+60',   flag: '🇲🇾', name: 'Malaysia'                           },
  { cc: '+960',  flag: '🇲🇻', name: 'Maldives'                           },
  { cc: '+223',  flag: '🇲🇱', name: 'Mali'                               },
  { cc: '+356',  flag: '🇲🇹', name: 'Malta'                              },
  { cc: '+692',  flag: '🇲🇭', name: 'Marshall Islands'                   },
  { cc: '+222',  flag: '🇲🇷', name: 'Mauritania'                         },
  { cc: '+230',  flag: '🇲🇺', name: 'Mauritius'                          },
  { cc: '+52',   flag: '🇲🇽', name: 'Mexico'                             },
  { cc: '+691',  flag: '🇫🇲', name: 'Micronesia'                         },
  { cc: '+373',  flag: '🇲🇩', name: 'Moldova'                            },
  { cc: '+377',  flag: '🇲🇨', name: 'Monaco'                             },
  { cc: '+976',  flag: '🇲🇳', name: 'Mongolia'                           },
  { cc: '+382',  flag: '🇲🇪', name: 'Montenegro'                         },
  { cc: '+212',  flag: '🇲🇦', name: 'Morocco'                            },
  { cc: '+258',  flag: '🇲🇿', name: 'Mozambique'                         },
  { cc: '+95',   flag: '🇲🇲', name: 'Myanmar'                            },
  { cc: '+264',  flag: '🇳🇦', name: 'Namibia'                            },
  { cc: '+674',  flag: '🇳🇷', name: 'Nauru'                              },
  { cc: '+977',  flag: '🇳🇵', name: 'Nepal'                              },
  { cc: '+31',   flag: '🇳🇱', name: 'Netherlands'                        },
  { cc: '+64',   flag: '🇳🇿', name: 'New Zealand'                        },
  { cc: '+505',  flag: '🇳🇮', name: 'Nicaragua'                          },
  { cc: '+227',  flag: '🇳🇪', name: 'Niger'                              },
  { cc: '+234',  flag: '🇳🇬', name: 'Nigeria'                            },
  { cc: '+850',  flag: '🇰🇵', name: 'North Korea'                        },
  { cc: '+389',  flag: '🇲🇰', name: 'North Macedonia'                    },
  { cc: '+47',   flag: '🇳🇴', name: 'Norway'                             },
  { cc: '+968',  flag: '🇴🇲', name: 'Oman'                               },
  { cc: '+92',   flag: '🇵🇰', name: 'Pakistan'                           },
  { cc: '+680',  flag: '🇵🇼', name: 'Palau'                              },
  { cc: '+970',  flag: '🇵🇸', name: 'Palestine'                          },
  { cc: '+507',  flag: '🇵🇦', name: 'Panama'                             },
  { cc: '+675',  flag: '🇵🇬', name: 'Papua New Guinea'                   },
  { cc: '+595',  flag: '🇵🇾', name: 'Paraguay'                           },
  { cc: '+51',   flag: '🇵🇪', name: 'Peru'                               },
  { cc: '+63',   flag: '🇵🇭', name: 'Philippines'                        },
  { cc: '+48',   flag: '🇵🇱', name: 'Poland'                             },
  { cc: '+351',  flag: '🇵🇹', name: 'Portugal'                           },
  { cc: '+974',  flag: '🇶🇦', name: 'Qatar'                              },
  { cc: '+40',   flag: '🇷🇴', name: 'Romania'                            },
  { cc: '+7',    flag: '🇷🇺', name: 'Russia'                             },
  { cc: '+250',  flag: '🇷🇼', name: 'Rwanda'                             },
  { cc: '+1869', flag: '🇰🇳', name: 'Saint Kitts and Nevis'              },
  { cc: '+1758', flag: '🇱🇨', name: 'Saint Lucia'                        },
  { cc: '+1784', flag: '🇻🇨', name: 'Saint Vincent and the Grenadines'   },
  { cc: '+685',  flag: '🇼🇸', name: 'Samoa'                              },
  { cc: '+378',  flag: '🇸🇲', name: 'San Marino'                         },
  { cc: '+239',  flag: '🇸🇹', name: 'Sao Tome and Principe'              },
  { cc: '+966',  flag: '🇸🇦', name: 'Saudi Arabia'                       },
  { cc: '+221',  flag: '🇸🇳', name: 'Senegal'                            },
  { cc: '+381',  flag: '🇷🇸', name: 'Serbia'                             },
  { cc: '+248',  flag: '🇸🇨', name: 'Seychelles'                         },
  { cc: '+232',  flag: '🇸🇱', name: 'Sierra Leone'                       },
  { cc: '+65',   flag: '🇸🇬', name: 'Singapore'                          },
  { cc: '+421',  flag: '🇸🇰', name: 'Slovakia'                           },
  { cc: '+386',  flag: '🇸🇮', name: 'Slovenia'                           },
  { cc: '+677',  flag: '🇸🇧', name: 'Solomon Islands'                    },
  { cc: '+252',  flag: '🇸🇴', name: 'Somalia'                            },
  { cc: '+27',   flag: '🇿🇦', name: 'South Africa'                       },
  { cc: '+82',   flag: '🇰🇷', name: 'South Korea'                        },
  { cc: '+211',  flag: '🇸🇸', name: 'South Sudan'                        },
  { cc: '+34',   flag: '🇪🇸', name: 'Spain'                              },
  { cc: '+94',   flag: '🇱🇰', name: 'Sri Lanka'                          },
  { cc: '+249',  flag: '🇸🇩', name: 'Sudan'                              },
  { cc: '+597',  flag: '🇸🇷', name: 'Suriname'                           },
  { cc: '+46',   flag: '🇸🇪', name: 'Sweden'                             },
  { cc: '+41',   flag: '🇨🇭', name: 'Switzerland'                        },
  { cc: '+963',  flag: '🇸🇾', name: 'Syria'                              },
  { cc: '+886',  flag: '🇹🇼', name: 'Taiwan'                             },
  { cc: '+992',  flag: '🇹🇯', name: 'Tajikistan'                         },
  { cc: '+255',  flag: '🇹🇿', name: 'Tanzania'                           },
  { cc: '+66',   flag: '🇹🇭', name: 'Thailand'                           },
  { cc: '+670',  flag: '🇹🇱', name: 'Timor-Leste'                        },
  { cc: '+228',  flag: '🇹🇬', name: 'Togo'                               },
  { cc: '+676',  flag: '🇹🇴', name: 'Tonga'                              },
  { cc: '+1868', flag: '🇹🇹', name: 'Trinidad and Tobago'                },
  { cc: '+216',  flag: '🇹🇳', name: 'Tunisia'                            },
  { cc: '+90',   flag: '🇹🇷', name: 'Turkey'                             },
  { cc: '+993',  flag: '🇹🇲', name: 'Turkmenistan'                       },
  { cc: '+688',  flag: '🇹🇻', name: 'Tuvalu'                             },
  { cc: '+256',  flag: '🇺🇬', name: 'Uganda'                             },
  { cc: '+380',  flag: '🇺🇦', name: 'Ukraine'                            },
  { cc: '+971',  flag: '🇦🇪', name: 'United Arab Emirates'               },
  { cc: '+44',   flag: '🇬🇧', name: 'United Kingdom'                     },
  { cc: '+1',    flag: '🇺🇸', name: 'United States'                      },
  { cc: '+598',  flag: '🇺🇾', name: 'Uruguay'                            },
  { cc: '+998',  flag: '🇺🇿', name: 'Uzbekistan'                         },
  { cc: '+678',  flag: '🇻🇺', name: 'Vanuatu'                            },
  { cc: '+379',  flag: '🇻🇦', name: 'Vatican City'                       },
  { cc: '+58',   flag: '🇻🇪', name: 'Venezuela'                          },
  { cc: '+84',   flag: '🇻🇳', name: 'Vietnam'                            },
  { cc: '+967',  flag: '🇾🇪', name: 'Yemen'                              },
  { cc: '+260',  flag: '🇿🇲', name: 'Zambia'                             },
  { cc: '+263',  flag: '🇿🇼', name: 'Zimbabwe'                           },
];
const DEFAULT_CC = '+61';
const FF           = 'Montserrat,sans-serif';
const LS_PERSONS   = '3edge-crm-persons';
const LS_COMPANIES = '3edge-crm-companies';

// ── Light-mode palette ────────────────────────────────────────────
const C = {
  bg:       '#f7f8fa',
  surface:  '#ffffff',
  border:   '#e2e5ea',
  borderMd: '#cdd1d9',
  text:     '#1a2030',
  sub:      '#4a5568',
  muted:    '#8a97a8',
  green:    '#2a9e2a',
  greenBg:  'rgba(42,158,42,.09)',
  greenBd:  'rgba(42,158,42,.35)',
  purple:   '#6c3fbf',
  red:      '#c0392b',
  thBg:     '#f0f2f6',
  rowHover: '#f5f7fb',
  tag:      '#eef0f5',
  tagText:  '#5a6a80',
};

// ── Helpers ───────────────────────────────────────────────────────
const uid      = (): string    => `${Date.now()}-${Math.random().toString(36).substr(2, 5)}`;

const ACTIVITY_TYPES: CrmActivityType[] = ['Phone call', 'Email', 'Text', 'In person'];

const todayIso = (): string => {
  const t = new Date();
  return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`;
};

const emptyActivity = (): CrmActivity => ({
  id: uid(),
  date: todayIso(),
  type: 'Phone call',
  notes: '',
  followUpDate: '',
  done: false,
});

const normalizePerson = (p: CrmPerson): CrmPerson => ({ ...p, activities: p.activities ?? [] });

const loadLS   = <T,>(k: string, fb: T): T => { try { const v = localStorage.getItem(k); return v ? (JSON.parse(v) as T) : fb; } catch { return fb; } };
const firstPhone = (arr: CrmPhone[]): string => { const p = arr.find(x => x.value); return p ? `${p.cc || DEFAULT_CC} ${p.value}` : '—'; };
const firstEmail = (arr: CrmEmail[]): string => arr.find(x => x.value)?.value || '—';

const mapsUrl = (address: string): string =>
  `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(address)}`;

const emptyPhone   = (): CrmPhone   => ({ value: '', type: 'Work', cc: DEFAULT_CC });
const emptyPerson  = (): CrmPerson  => ({
  id: uid(), name: '', organizationId: '', position: '',
  phones: [emptyPhone()], emails: [{ value: '', type: 'Work' }], activities: [],
});
const emptyCompany = (): CrmCompany => ({ id: uid(), name: '', labels: '', address: '', phones: [emptyPhone()], emails: [{ value: '', type: 'Work' }] });

// ── Shared modal input style ──────────────────────────────────────
const mi: React.CSSProperties = {
  padding: '8px 10px', background: C.surface, border: `1px solid ${C.borderMd}`,
  borderRadius: 4, color: C.text, fontSize: 12.5, fontFamily: FF,
  width: '100%', boxSizing: 'border-box', outline: 'none',
};
const ml: React.CSSProperties = {
  fontSize: 10, fontWeight: 700, color: C.sub, letterSpacing: '.07em',
  textTransform: 'uppercase', marginBottom: 4, display: 'block', fontFamily: FF,
};

// ── Address autocomplete ──────────────────────────────────────────
interface NominatimResult { display_name: string; }

const AddressSearch: React.FC<{ value: string; onChange: (v: string) => void }> = ({ value, onChange }) => {
  const [suggestions, setSuggestions] = React.useState<string[]>([]);
  const [open, setOpen]               = React.useState(false);
  const timer = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const wrapRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const handler = (e: MouseEvent): void => {
      if (wrapRef.current && !wrapRef.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const handleChange = (q: string): void => {
    onChange(q);
    if (timer.current) clearTimeout(timer.current);
    if (q.length < 3) { setSuggestions([]); setOpen(false); return; }
    timer.current = setTimeout(async () => {
      try {
        const res  = await fetch(`https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(q)}&format=json&limit=6`, { headers: { 'Accept-Language': 'en' } });
        const data = await res.json() as NominatimResult[];
        setSuggestions(data.map(d => d.display_name));
        setOpen(true);
      } catch { setSuggestions([]); }
    }, 420);
  };

  const pick = (s: string): void => { onChange(s); setSuggestions([]); setOpen(false); };

  return (
    <div ref={wrapRef} style={{ position: 'relative' }}>
      <input
        value={value}
        onChange={e => handleChange(e.target.value)}
        onFocus={() => suggestions.length > 0 && setOpen(true)}
        placeholder="Start typing an address…"
        style={mi}
      />
      {open && suggestions.length > 0 && (
        <div style={{ position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0, background: C.surface, border: `1px solid ${C.borderMd}`, borderRadius: 6, zIndex: 200, boxShadow: '0 6px 20px rgba(0,0,0,.12)', overflow: 'hidden' }}>
          {suggestions.map((s, i) => (
            <div
              key={i}
              onMouseDown={() => pick(s)}
              style={{ padding: '9px 12px', fontSize: 12, fontFamily: FF, color: C.text, cursor: 'pointer', borderBottom: i < suggestions.length - 1 ? `1px solid ${C.border}` : 'none', display: 'flex', alignItems: 'flex-start', gap: 8 }}
              onMouseEnter={e => { (e.currentTarget as HTMLDivElement).style.background = C.rowHover; }}
              onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = 'transparent'; }}
            >
              <svg style={{ flexShrink: 0, marginTop: 1 }} width="13" height="13" viewBox="0 0 24 24" fill="none" stroke={C.muted} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/>
              </svg>
              <span style={{ lineHeight: 1.4 }}>{s}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

// ── Modal wrapper (light) ─────────────────────────────────────────
const Modal: React.FC<{ title: string; onClose: () => void; children: React.ReactNode }> = ({ title, onClose, children }) => (
  <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.45)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
    <div style={{ background: C.surface, borderRadius: 8, width: 500, maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 12px 40px rgba(0,0,0,.18)', border: `1px solid ${C.border}` }}>
      <div style={{ padding: '14px 20px', borderBottom: `1px solid ${C.border}`, display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: C.thBg, borderRadius: '8px 8px 0 0' }}>
        <span style={{ fontFamily: FF, fontWeight: 700, fontSize: 13, color: C.text, letterSpacing: '.04em' }}>{title}</span>
        <button onClick={onClose} style={{ background: 'none', border: 'none', color: C.muted, cursor: 'pointer', fontSize: 20, lineHeight: 1, padding: 0 }}>×</button>
      </div>
      <div style={{ padding: '20px 22px' }}>{children}</div>
    </div>
  </div>
);

// ── Auto-detect country code from a pasted/typed number ──────────
const CC_SORTED = [...COUNTRY_CODES].sort((a, b) => b.cc.length - a.cc.length);

const detectAndSplit = (raw: string): { cc: string; num: string } | undefined => {
  const stripped = raw.replace(/[\s\-().]/g, '');
  if (stripped.length < 4) return undefined;
  const withPlus = stripped.startsWith('+') ? stripped : '+' + stripped;
  for (const c of CC_SORTED) {
    if (withPlus.startsWith(c.cc) && withPlus.length > c.cc.length) {
      return { cc: c.cc, num: withPlus.slice(c.cc.length) };
    }
  }
  return undefined;
};

// ── Multi-value field (phone / email) ─────────────────────────────
const MultiField: React.FC<{
  items: CrmPhone[];
  types: string[];
  addLabel: string;
  placeholder: string;
  isPhone?: boolean;
  onChange: (items: CrmPhone[]) => void;
}> = ({ items, types, addLabel, placeholder, isPhone, onChange }) => (
  <div style={{ marginBottom: 14 }}>
    {items.map((item, i) => (
      <div key={i} style={{ display: 'flex', gap: 5, marginBottom: 6, alignItems: 'center' }}>
        {isPhone && (
          <select
            value={item.cc || DEFAULT_CC}
            onChange={e => { const n = [...items]; n[i] = { ...n[i], cc: e.target.value }; onChange(n); }}
            style={{ ...mi, width: 175, flex: 'none', paddingLeft: 6, paddingRight: 4, fontSize: 12 }}
          >
            {COUNTRY_CODES.map(c => (
              <option key={c.cc + c.name} value={c.cc}>{c.flag} {c.name} ({c.cc})</option>
            ))}
          </select>
        )}
        <input
          value={item.value}
          onChange={e => {
            const raw = e.target.value;
            const detected = isPhone ? detectAndSplit(raw) : null;
            const n = [...items];
            n[i] = detected ? { ...n[i], cc: detected.cc, value: detected.num } : { ...n[i], value: raw };
            onChange(n);
          }}
          placeholder={placeholder}
          autoComplete="off"
          style={{ ...mi, flex: 1 }}
        />
        <select value={item.type} onChange={e => { const n = [...items]; n[i] = { ...n[i], type: e.target.value }; onChange(n); }} style={{ ...mi, width: 86, flex: 'none' }}>
          {types.map(t => <option key={t} value={t}>{t}</option>)}
        </select>
        {items.length > 1 && <button onClick={() => onChange(items.filter((_, j) => j !== i))} style={{ background: 'none', border: 'none', color: C.muted, cursor: 'pointer', fontSize: 18, lineHeight: 1, padding: '0 2px', flexShrink: 0 }}>×</button>}
      </div>
    ))}
    <button onClick={() => onChange([...items, isPhone ? emptyPhone() : { value: '', type: types[0] }])} style={{ background: 'none', border: 'none', color: C.green, fontFamily: FF, fontSize: 12, cursor: 'pointer', padding: 0 }}>
      + Add {addLabel}
    </button>
  </div>
);

// ── Modal footer ──────────────────────────────────────────────────
const ModalFooter: React.FC<{ onCancel: () => void; onSave: () => void }> = ({ onCancel, onSave }) => (
  <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 20, paddingTop: 14, borderTop: `1px solid ${C.border}` }}>
    <button onClick={onCancel} style={{ padding: '8px 20px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: 'transparent', color: C.sub, fontFamily: FF, fontSize: 12, cursor: 'pointer' }}>Cancel</button>
    <button onClick={onSave}   style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: C.green, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: 'pointer' }}>Save</button>
  </div>
);

// ── Person Modal ──────────────────────────────────────────────────
const PersonModal: React.FC<{ initial: CrmPerson; companies: CrmCompany[]; onSave: (p: CrmPerson) => void; onClose: () => void }> = ({ initial, companies, onSave, onClose }) => {
  const [d, setD] = React.useState<CrmPerson>(() => normalizePerson(initial));
  const [draft, setDraft] = React.useState<CrmActivity>(() => emptyActivity());
  const set = <K extends keyof CrmPerson>(k: K, v: CrmPerson[K]): void => setD(p => ({ ...p, [k]: v }));

  React.useEffect(() => {
    setD(normalizePerson(initial));
    setDraft(emptyActivity());
  }, [initial.id]);

  const setDraftField = <K extends keyof CrmActivity>(k: K, v: CrmActivity[K]): void =>
    setDraft(a => ({ ...a, [k]: v }));

  const addActivity = (): void => {
    setD(p => ({ ...p, activities: [{ ...draft, id: uid() }, ...(p.activities ?? [])] }));
    setDraft(emptyActivity());
  };

  const removeActivity = (id: string): void => {
    setD(p => ({ ...p, activities: (p.activities ?? []).filter(a => a.id !== id) }));
  };

  const row2: React.CSSProperties = { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 };

  return (
    <Modal title={initial.name ? 'Edit Person' : 'Add Person'} onClose={onClose}>
      <div style={{ marginBottom: 14 }}>
        <label style={ml}>Name</label>
        <input value={d.name} onChange={e => set('name', e.target.value)} style={mi} placeholder="Full name" autoFocus />
      </div>
      <div style={{ marginBottom: 14 }}>
        <label style={ml}>Company</label>
        <div style={{ position: 'relative' }}>
          <span style={{ position: 'absolute', left: 9, top: '50%', transform: 'translateY(-50%)', color: C.muted, pointerEvents: 'none' }}>
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 7V5a2 2 0 0 0-4 0v2"/></svg>
          </span>
          <select value={d.organizationId} onChange={e => set('organizationId', e.target.value)} style={{ ...mi, paddingLeft: 28 }}>
            <option value="">— None —</option>
            {companies.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
        </div>
      </div>
      <div style={{ marginBottom: 14 }}>
        <label style={ml}>Position</label>
        <select value={d.position} onChange={e => set('position', e.target.value)} style={mi}>
          <option value="">— Select position —</option>
          <option>Project Manager</option>
          <option>Estimator</option>
          <option>Admin</option>
          <option>Accounts</option>
          <option>Construction Mgr</option>
          <option>Owner</option>
          <option>Engineer</option>
          <option>Architect</option>
        </select>
      </div>
      <label style={ml}>Phone</label>
      <MultiField items={d.phones} types={PHONE_TYPES} addLabel="phone" placeholder="Phone number" isPhone onChange={v => set('phones', v)} />
      <label style={ml}>Email</label>
      <MultiField items={d.emails} types={EMAIL_TYPES} addLabel="email" placeholder="Email address" onChange={v => set('emails', v)} />

      <div style={{ marginTop: 20, paddingTop: 16, borderTop: `1px solid ${C.border}` }}>
        <div style={{ fontFamily: FF, fontWeight: 700, fontSize: 11.5, letterSpacing: '.06em', textTransform: 'uppercase', color: C.text, marginBottom: 12 }}>
          Activity Log
        </div>
        <div style={row2}>
          <div>
            <label style={ml}>Date</label>
            <input type="date" value={draft.date} onChange={e => setDraftField('date', e.target.value)} style={mi} />
          </div>
          <div>
            <label style={ml}>Type</label>
            <select value={draft.type} onChange={e => setDraftField('type', e.target.value as CrmActivityType)} style={mi}>
              {ACTIVITY_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
            </select>
          </div>
        </div>
        <div style={{ marginTop: 10 }}>
          <label style={ml}>Notes</label>
          <textarea
            value={draft.notes}
            onChange={e => setDraftField('notes', e.target.value)}
            placeholder="What was discussed…"
            rows={3}
            style={{ ...mi, resize: 'vertical', minHeight: 64 }}
          />
        </div>
        <div style={{ ...row2, marginTop: 10, alignItems: 'end' }}>
          <div>
            <label style={ml}>Follow up date</label>
            <input type="date" value={draft.followUpDate} onChange={e => setDraftField('followUpDate', e.target.value)} style={mi} />
          </div>
          <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontFamily: FF, fontSize: 12.5, color: C.text, cursor: 'pointer', paddingBottom: 8 }}>
            <input
              type="checkbox"
              checked={draft.done}
              onChange={e => setDraftField('done', e.target.checked)}
              style={{ width: 16, height: 16, accentColor: C.green, cursor: 'pointer' }}
            />
            Mark as done
          </label>
        </div>
        <button
          type="button"
          onClick={addActivity}
          style={{ marginTop: 10, padding: '6px 14px', borderRadius: 4, border: `1px solid ${C.greenBd}`, background: C.greenBg, color: C.green, fontFamily: FF, fontWeight: 700, fontSize: 11.5, cursor: 'pointer' }}
        >
          + Add log entry
        </button>

        {(d.activities ?? []).length > 0 && (
          <div style={{ marginTop: 14, maxHeight: 160, overflowY: 'auto', border: `1px solid ${C.border}`, borderRadius: 4 }}>
            {(d.activities ?? []).map(a => (
              <div
                key={a.id}
                style={{
                  padding: '10px 12px', borderBottom: `1px solid ${C.border}`, fontFamily: FF, fontSize: 12,
                  opacity: a.done ? 0.65 : 1, background: a.done ? C.thBg : C.surface,
                }}
              >
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 8 }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 700, color: C.text, textDecoration: a.done ? 'line-through' : 'none' }}>
                      {a.date} · {a.type}{a.done ? ' ✓' : ''}
                    </div>
                    {a.notes && <div style={{ color: C.sub, marginTop: 4, lineHeight: 1.4, whiteSpace: 'pre-wrap' }}>{a.notes}</div>}
                    {a.followUpDate && (
                      <div style={{ color: C.muted, marginTop: 4, fontSize: 11 }}>
                        Follow up: {a.followUpDate}
                      </div>
                    )}
                  </div>
                  <button
                    type="button"
                    onClick={() => removeActivity(a.id)}
                    style={{ background: 'none', border: 'none', color: C.muted, cursor: 'pointer', fontSize: 16, lineHeight: 1, padding: 0 }}
                    title="Remove entry"
                  >
                    ×
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      <ModalFooter onCancel={onClose} onSave={() => { if (d.name.trim()) onSave(normalizePerson(d)); }} />
    </Modal>
  );
};

// ── CRM Import Modal ──────────────────────────────────────────────
const CrmImportModal: React.FC<{
  onClose: () => void;
  onImport: (orgBuf: ArrayBuffer | null, peopleBuf: ArrayBuffer | null) => void;
}> = ({ onClose, onImport }) => {
  const [orgBuf, setOrgBuf]       = React.useState<ArrayBuffer | null>(null);
  const [peopleBuf, setPeopleBuf] = React.useState<ArrayBuffer | null>(null);
  const [orgName, setOrgName]     = React.useState('');
  const [peopleName, setPeopleName] = React.useState('');
  const [error, setError]         = React.useState('');

  const readFile = (f: File | null, setter: (b: ArrayBuffer | null) => void, nameSetter: (n: string) => void): void => {
    if (!f) return;
    setError('');
    const reader = new FileReader();
    reader.onload = ev => { setter(ev.target!.result as ArrayBuffer); nameSetter(f.name); };
    reader.onerror = () => setError('Failed to read file.');
    reader.readAsArrayBuffer(f);
  };

  return (
    <Modal title="Import CRM Data" onClose={onClose}>
      <p style={{ fontFamily: FF, fontSize: 12, color: C.sub, lineHeight: 1.55, margin: '0 0 16px 0' }}>
        Upload Pipedrive exports. Each file replaces that list entirely and imports every row
        (645 people, 512 organization rows — every Excel row is kept, even duplicate names).
        Import organizations before people so company links work.
      </p>
      <div style={{ marginBottom: 14 }}>
        <label style={ml}>Organizations (.xlsx)</label>
        <input
          type="file"
          accept=".xls,.xlsx"
          onChange={e => readFile(e.target.files?.[0] || null, setOrgBuf, setOrgName)}
          style={{ fontFamily: FF, fontSize: 12, width: '100%' }}
        />
        {orgName && <div style={{ fontSize: 11, color: C.green, marginTop: 4, fontFamily: FF }}>{orgName}</div>}
      </div>
      <div style={{ marginBottom: 14 }}>
        <label style={ml}>People (.xlsx)</label>
        <input
          type="file"
          accept=".xls,.xlsx"
          onChange={e => readFile(e.target.files?.[0] || null, setPeopleBuf, setPeopleName)}
          style={{ fontFamily: FF, fontSize: 12, width: '100%' }}
        />
        {peopleName && <div style={{ fontSize: 11, color: C.green, marginTop: 4, fontFamily: FF }}>{peopleName}</div>}
      </div>
      {error && (
        <div style={{ padding: '10px 12px', borderRadius: 4, background: 'rgba(192,57,43,.08)', border: `1px solid ${C.red}`, color: C.red, fontFamily: FF, fontSize: 12, marginBottom: 12 }}>
          {error}
        </div>
      )}
      <ModalFooter
        onCancel={onClose}
        onSave={() => {
          if (!orgBuf && !peopleBuf) { setError('Select at least one file.'); return; }
          try {
            onImport(orgBuf, peopleBuf);
          } catch (e) {
            setError(String(e instanceof Error ? e.message : e));
          }
        }}
      />
    </Modal>
  );
};

// ── Company Modal ─────────────────────────────────────────────────
const CompanyModal: React.FC<{ initial: CrmCompany; onSave: (c: CrmCompany) => void; onClose: () => void }> = ({ initial, onSave, onClose }) => {
  const [d, setD] = React.useState<CrmCompany>(initial);
  const set = <K extends keyof CrmCompany>(k: K, v: CrmCompany[K]): void => setD(p => ({ ...p, [k]: v }));
  return (
    <Modal title={initial.name ? 'Edit Company' : 'Add Company'} onClose={onClose}>
      <div style={{ marginBottom: 14 }}>
        <label style={ml}>Name</label>
        <input value={d.name} onChange={e => set('name', e.target.value)} style={mi} placeholder="Company name" autoFocus />
      </div>
      <div style={{ marginBottom: 14 }}>
        <label style={ml}>Labels</label>
        <input value={d.labels} onChange={e => set('labels', e.target.value)} style={mi} placeholder="e.g. Client, Supplier, Partner" />
      </div>
      <div style={{ marginBottom: 14 }}>
        <label style={ml}>Address</label>
        <AddressSearch value={d.address} onChange={v => set('address', v)} />
      </div>
      <label style={ml}>Phone</label>
      <MultiField items={d.phones} types={PHONE_TYPES} addLabel="phone" placeholder="Phone number" isPhone onChange={v => set('phones', v)} />
      <label style={ml}>Email</label>
      <MultiField items={d.emails} types={EMAIL_TYPES} addLabel="email" placeholder="Email address" onChange={v => set('emails', v)} />
      <ModalFooter onCancel={onClose} onSave={() => { if (d.name.trim()) onSave(d); }} />
    </Modal>
  );
};

// ── Table helpers ─────────────────────────────────────────────────
const tableStyle: React.CSSProperties = {
  width: '100%',
  borderCollapse: 'collapse',
  tableLayout: 'fixed',
};

const thStyle = (opts?: { w?: string | number; center?: boolean }): React.CSSProperties => ({
  padding: '9px 8px',
  textAlign: opts?.center ? 'center' : 'left',
  fontFamily: FF,
  fontWeight: 700,
  fontSize: 10.5,
  letterSpacing: '.07em',
  textTransform: 'uppercase',
  color: C.sub,
  background: C.thBg,
  borderBottom: `2px solid ${C.borderMd}`,
  whiteSpace: 'normal',
  lineHeight: 1.3,
  width: opts?.w,
});

const Th: React.FC<{ label: string; w?: string | number; center?: boolean }> = ({ label, w, center }) => (
  <th style={thStyle({ w, center })}>{label}</th>
);

type SortDir = 'asc' | 'desc';
type PersonSortKey = 'name' | 'company' | 'position' | 'location' | 'phone' | 'email';
type CompanySortKey = 'name' | 'labels' | 'phone' | 'email' | 'address';

const cmpAlpha = (a: string, b: string, dir: SortDir): number => {
  const r = a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true });
  return dir === 'asc' ? r : -r;
};

const SortTh: React.FC<{
  label: string;
  active: boolean;
  dir: SortDir;
  onClick: () => void;
  center?: boolean;
}> = ({ label, active, dir, onClick, center }) => (
  <th
    onClick={onClick}
    title="Click to sort A–Z"
    style={{
      ...thStyle({ center }),
      cursor: 'pointer',
      userSelect: 'none',
      color: active ? C.green : C.sub,
    }}
  >
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
      {label}
      <span style={{ fontSize: 9, lineHeight: 1, opacity: active ? 1 : 0.4 }} aria-hidden>
        {active ? (dir === 'asc' ? '▲' : '▼') : '↕'}
      </span>
    </span>
  </th>
);

const tdBase = (muted?: boolean, mono?: boolean): React.CSSProperties => ({
  padding: '8px 8px',
  fontFamily: mono ? 'monospace' : FF,
  fontSize: 12.5,
  color: muted ? C.muted : C.text,
  borderBottom: `1px solid ${C.border}`,
  verticalAlign: 'middle',
  wordBreak: 'break-word',
});

/** Up to 2 lines of text, then ellipsis */
const cellWrap: React.CSSProperties = {
  display: '-webkit-box',
  WebkitLineClamp: 2,
  WebkitBoxOrient: 'vertical',
  overflow: 'hidden',
  whiteSpace: 'normal',
  wordBreak: 'break-word',
  lineHeight: 1.35,
};

const Td: React.FC<{ children: React.ReactNode; muted?: boolean; mono?: boolean; wrap?: boolean }> = ({ children, muted, mono, wrap = true }) => (
  <td style={tdBase(muted, mono)}>
    {wrap ? <div style={cellWrap}>{children}</div> : children}
  </td>
);

const TdIndex: React.FC<{ children: React.ReactNode }> = ({ children }) => (
  <td style={{ ...tdBase(true), textAlign: 'center', whiteSpace: 'nowrap', padding: '8px 6px' }}>{children}</td>
);

// ── CrmBoard ──────────────────────────────────────────────────────
const CrmBoard: React.FC = () => {
  const [tab, setTab]             = React.useState<'persons' | 'companies'>('persons');
  const [persons, setPersons]     = React.useState<CrmPerson[]>(() => loadLS<CrmPerson[]>(LS_PERSONS, []));
  const [companies, setCompanies] = React.useState<CrmCompany[]>(() => loadLS<CrmCompany[]>(LS_COMPANIES, []));
  const [personModal, setPersonModal]   = React.useState<CrmPerson | null>(null);
  const [companyModal, setCompanyModal] = React.useState<CrmCompany | null>(null);
  const [importModal, setImportModal]   = React.useState(false);
  const [search, setSearch]             = React.useState('');
  const [personSort, setPersonSort]     = React.useState<{ key: PersonSortKey; dir: SortDir } | null>(null);
  const [companySort, setCompanySort] = React.useState<{ key: CompanySortKey; dir: SortDir } | null>(null);

  React.useEffect(() => { localStorage.setItem(LS_PERSONS,   JSON.stringify(persons));   }, [persons]);
  React.useEffect(() => { localStorage.setItem(LS_COMPANIES, JSON.stringify(companies)); }, [companies]);

  const savePerson    = (p: CrmPerson):  void => { setPersons(prev  => prev.some(x => x.id === p.id) ? prev.map(x => x.id === p.id ? p : x) : [...prev, p]);   setPersonModal(null); };
  const deletePerson  = (id: string):    void => { if (confirm('Delete this person?'))  setPersons(prev  => prev.filter(x => x.id !== id)); };
  const saveCompany   = (c: CrmCompany): void => { setCompanies(prev => prev.some(x => x.id === c.id) ? prev.map(x => x.id === c.id ? c : x) : [...prev, c]); setCompanyModal(null); };
  const deleteCompany = (id: string):    void => { if (confirm('Delete this company?')) setCompanies(prev => prev.filter(x => x.id !== id)); };
  const companyName   = (id: string):    string => companies.find(c => c.id === id)?.name || '';
  const companyAddress = (id: string):   string => companies.find(c => c.id === id)?.address || '';

  const handleImport = (orgBuf: ArrayBuffer | null, peopleBuf: ArrayBuffer | null): void => {
    const { companies: nextCo, persons: nextPe } = runCrmImport(companies, persons, orgBuf, peopleBuf);
    setCompanies(nextCo);
    setPersons(nextPe);
    setImportModal(false);
  };

  const q = search.toLowerCase();
  const visPersons  = persons.filter(p  => !q || p.name.toLowerCase().includes(q) || companyName(p.organizationId).toLowerCase().includes(q) || companyAddress(p.organizationId).toLowerCase().includes(q));
  const visCompanies = companies.filter(c => !q || c.name.toLowerCase().includes(q) || c.labels.toLowerCase().includes(q) || c.address.toLowerCase().includes(q));

  const togglePersonSort = (key: PersonSortKey): void => {
    setPersonSort(prev => (prev?.key === key
      ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' }
      : { key, dir: 'asc' }));
  };

  const toggleCompanySort = (key: CompanySortKey): void => {
    setCompanySort(prev => (prev?.key === key
      ? { key, dir: prev.dir === 'asc' ? 'desc' : 'asc' }
      : { key, dir: 'asc' }));
  };

  const personSortVal = (p: CrmPerson, key: PersonSortKey): string => {
    switch (key) {
      case 'name': return p.name;
      case 'company': return companyName(p.organizationId);
      case 'position': return p.position;
      case 'location': return companyAddress(p.organizationId);
      case 'phone': return firstPhone(p.phones);
      case 'email': return firstEmail(p.emails);
      default: return '';
    }
  };

  const sortedPersons = React.useMemo(() => {
    const list = [...visPersons];
    if (!personSort) return list;
    const { key, dir } = personSort;
    list.sort((a, b) => cmpAlpha(personSortVal(a, key).toLowerCase(), personSortVal(b, key).toLowerCase(), dir));
    return list;
  }, [visPersons, personSort, companies]);

  const sortedCompanies = React.useMemo(() => {
    const list = [...visCompanies];
    if (!companySort) return list;
    const { key, dir } = companySort;
    list.sort((a, b) => {
      let av = '';
      let bv = '';
      switch (key) {
        case 'name': av = a.name; bv = b.name; break;
        case 'labels': av = a.labels; bv = b.labels; break;
        case 'phone': av = firstPhone(a.phones); bv = firstPhone(b.phones); break;
        case 'email': av = firstEmail(a.emails); bv = firstEmail(b.emails); break;
        case 'address': av = a.address; bv = b.address; break;
        default: break;
      }
      return cmpAlpha(av.toLowerCase(), bv.toLowerCase(), dir);
    });
    return list;
  }, [visCompanies, companySort]);

  const tabBtn = (active: boolean): React.CSSProperties => ({
    fontFamily: FF, fontWeight: 700, fontSize: 11.5, letterSpacing: '.07em', textTransform: 'uppercase',
    padding: '7px 22px', borderRadius: '4px 4px 0 0', cursor: 'pointer', transition: 'all .12s',
    background: active ? C.surface : 'transparent',
    borderTop:    active ? `2px solid ${C.green}` : '2px solid transparent',
    borderLeft:   active ? `1px solid ${C.border}` : '1px solid transparent',
    borderRight:  active ? `1px solid ${C.border}` : '1px solid transparent',
    borderBottom: active ? `1px solid ${C.surface}` : '1px solid transparent',
    color: active ? C.green : C.muted,
    marginBottom: active ? -1 : 0,
  });

  const actionBtn: React.CSSProperties = {
    padding: '4px 10px',
    borderRadius: 3,
    fontFamily: FF,
    fontWeight: 700,
    fontSize: 10.5,
    cursor: 'pointer',
    width: '100%',
    boxSizing: 'border-box',
    textAlign: 'center',
  };

  const ActionCell: React.FC<{ onEdit: () => void; onDel: () => void }> = ({ onEdit, onDel }) => (
    <td style={{ padding: '8px 10px 8px 6px', borderBottom: `1px solid ${C.border}`, verticalAlign: 'middle' }}>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 4, minWidth: 52 }}>
        <button onClick={onEdit} style={{ ...actionBtn, border: 'none', background: C.purple, color: '#fff' }}>Edit</button>
        <button onClick={onDel}  style={{ ...actionBtn, border: `1px solid ${C.red}`, background: 'transparent', color: C.red }}>Del</button>
      </div>
    </td>
  );

  const emptyRow = (colspan: number, msg: string): React.ReactNode => (
    <tr><td colSpan={colspan} style={{ padding: '40px 0', textAlign: 'center', fontFamily: FF, fontSize: 13, color: C.muted, borderBottom: `1px solid ${C.border}` }}>{msg}</td></tr>
  );

  return (
    <div style={{ background: C.bg, minHeight: 400, borderRadius: 8, padding: '0 0 24px 0', width: '100%', maxWidth: '100%', boxSizing: 'border-box' }}>

      {/* ── Toolbar */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-end', marginBottom: 0, padding: '16px 0 0 0' }}>
        <div style={{ display: 'flex', gap: 0, borderBottom: `1px solid ${C.border}`, paddingBottom: 0 }}>
          <button style={tabBtn(tab === 'persons')}   onClick={() => { setTab('persons');   setSearch(''); setCompanySort(null); }}>Persons</button>
          <button style={tabBtn(tab === 'companies')} onClick={() => { setTab('companies'); setSearch(''); setPersonSort(null); }}>Companies</button>
        </div>
        <div style={{ display: 'flex', gap: 10, alignItems: 'center', paddingBottom: 4 }}>
          <input
            value={search}
            onChange={e => setSearch(e.target.value)}
            placeholder={`Search ${tab}…`}
            style={{ padding: '6px 10px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontSize: 12, color: C.text, outline: 'none', width: 200 }}
          />
          <button
            onClick={() => setImportModal(true)}
            style={{ padding: '7px 14px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, color: C.sub, fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: 'pointer', whiteSpace: 'nowrap' }}
          >
            Import
          </button>
          {tab === 'persons' ? (
            <button onClick={() => setPersonModal(emptyPerson())} style={{ padding: '7px 16px', borderRadius: 4, border: 'none', background: C.green, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: 'pointer', whiteSpace: 'nowrap' }}>+ Person</button>
          ) : (
            <button onClick={() => setCompanyModal(emptyCompany())} style={{ padding: '7px 16px', borderRadius: 4, border: 'none', background: C.green, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: 'pointer', whiteSpace: 'nowrap' }}>+ Company</button>
          )}
        </div>
      </div>

      {/* ── Table */}
      <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderTop: 'none', borderRadius: '0 0 8px 8px' }}>

        {/* Persons table */}
        {tab === 'persons' && (
          <table style={tableStyle}>
            <colgroup>
              <col style={{ width: '5%' }} />
              <col style={{ width: '11%' }} />
              <col style={{ width: '13%' }} />
              <col style={{ width: '9%' }} />
              <col style={{ width: '20%' }} />
              <col style={{ width: '12%' }} />
              <col style={{ width: '13%' }} />
              <col style={{ width: '9%' }} />
            </colgroup>
            <thead>
              <tr>
                <Th label="#" center />
                <SortTh label="Name"     active={personSort?.key === 'name'}     dir={personSort?.dir ?? 'asc'} onClick={() => togglePersonSort('name')} />
                <SortTh label="Company"  active={personSort?.key === 'company'}  dir={personSort?.dir ?? 'asc'} onClick={() => togglePersonSort('company')} />
                <SortTh label="Position" active={personSort?.key === 'position'} dir={personSort?.dir ?? 'asc'} onClick={() => togglePersonSort('position')} />
                <SortTh label="Location" active={personSort?.key === 'location'} dir={personSort?.dir ?? 'asc'} onClick={() => togglePersonSort('location')} />
                <SortTh label="Phone"    active={personSort?.key === 'phone'}    dir={personSort?.dir ?? 'asc'} onClick={() => togglePersonSort('phone')} />
                <SortTh label="Email"    active={personSort?.key === 'email'}    dir={personSort?.dir ?? 'asc'} onClick={() => togglePersonSort('email')} />
                <Th label="Actions"  />
              </tr>
            </thead>
            <tbody>
              {sortedPersons.length === 0
                ? emptyRow(8, persons.length === 0 ? 'No persons yet — click + Person or Import to add data.' : 'No results match your search.')
                : sortedPersons.map((p, idx) => (
                  <tr key={p.id}
                    onMouseEnter={e => { (e.currentTarget as HTMLTableRowElement).style.background = C.rowHover; }}
                    onMouseLeave={e => { (e.currentTarget as HTMLTableRowElement).style.background = 'transparent'; }}
                  >
                    <TdIndex>{idx + 1}</TdIndex>
                    <Td><span style={{ fontWeight: 700 }} title={p.name}>{p.name}</span></Td>
                    <Td>
                      {p.organizationId
                        ? <span style={{ color: C.green, fontWeight: 600 }} title={companyName(p.organizationId)}>{companyName(p.organizationId)}</span>
                        : <span style={{ color: C.muted }}>—</span>}
                    </Td>
                    <Td><span style={{ color: p.position ? C.text : C.muted, fontWeight: p.position ? 700 : 400 }} title={p.position}>{p.position || '—'}</span></Td>
                    <Td muted>
                      {(() => {
                        const addr = companyAddress(p.organizationId);
                        if (!addr) return '—';
                        return (
                          <a
                            href={mapsUrl(addr)}
                            target="_blank"
                            rel="noopener noreferrer"
                            title={addr}
                            style={{ color: C.green, fontWeight: 600, textDecoration: 'none' }}
                            onClick={e => e.stopPropagation()}
                          >
                            {addr}
                          </a>
                        );
                      })()}
                    </Td>
                    <Td><span style={{ color: C.sub, fontWeight: 600 }} title={firstPhone(p.phones)}>{firstPhone(p.phones)}</span></Td>
                    <Td><span style={{ color: C.sub, fontWeight: 600 }} title={firstEmail(p.emails)}>{firstEmail(p.emails)}</span></Td>
                    <ActionCell onEdit={() => setPersonModal(normalizePerson(p))} onDel={() => deletePerson(p.id)} />
                  </tr>
                ))
              }
            </tbody>
          </table>
        )}

        {/* Companies table */}
        {tab === 'companies' && (
          <table style={tableStyle}>
            <colgroup>
              <col style={{ width: '5%' }} />
              <col style={{ width: '17%' }} />
              <col style={{ width: '11%' }} />
              <col style={{ width: '12%' }} />
              <col style={{ width: '15%' }} />
              <col style={{ width: '23%' }} />
              <col style={{ width: '9%' }} />
            </colgroup>
            <thead>
              <tr>
                <Th label="#" center />
                <SortTh label="Company" active={companySort?.key === 'name'}    dir={companySort?.dir ?? 'asc'} onClick={() => toggleCompanySort('name')} />
                <SortTh label="Labels"  active={companySort?.key === 'labels'}  dir={companySort?.dir ?? 'asc'} onClick={() => toggleCompanySort('labels')} />
                <SortTh label="Phone"   active={companySort?.key === 'phone'}   dir={companySort?.dir ?? 'asc'} onClick={() => toggleCompanySort('phone')} />
                <SortTh label="Email"   active={companySort?.key === 'email'}   dir={companySort?.dir ?? 'asc'} onClick={() => toggleCompanySort('email')} />
                <SortTh label="Address" active={companySort?.key === 'address'} dir={companySort?.dir ?? 'asc'} onClick={() => toggleCompanySort('address')} />
                <Th label="Actions"  />
              </tr>
            </thead>
            <tbody>
              {sortedCompanies.length === 0
                ? emptyRow(7, companies.length === 0 ? 'No companies yet — click + Company to add one.' : 'No results match your search.')
                : sortedCompanies.map((c, idx) => {
                  const linked = persons.filter(p => p.organizationId === c.id);
                  return (
                    <tr key={c.id}
                      onMouseEnter={e => { (e.currentTarget as HTMLTableRowElement).style.background = C.rowHover; }}
                      onMouseLeave={e => { (e.currentTarget as HTMLTableRowElement).style.background = 'transparent'; }}
                    >
                      <TdIndex>{idx + 1}</TdIndex>
                      <Td>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
                          <span style={{ fontWeight: 700 }} title={c.name}>{c.name}</span>
                          {linked.length > 0 && (
                            <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
                              {linked.map(p => (
                                <span key={p.id} style={{ fontSize: 10, fontFamily: FF, color: C.green, background: C.greenBg, border: `1px solid ${C.greenBd}`, borderRadius: 3, padding: '1px 6px', cursor: 'pointer' }} onClick={() => { setTab('persons'); setSearch(p.name); }} title={companyAddress(c.id) || 'View person'}>
                                  {p.name}
                                </span>
                              ))}
                            </div>
                          )}
                        </div>
                      </Td>
                      <Td>
                        <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
                          {c.labels
                            ? c.labels.split(',').map(l => l.trim()).filter(Boolean).map(l => (
                                <span key={l} style={{ fontSize: 10, fontFamily: FF, color: C.tagText, background: C.tag, border: `1px solid ${C.border}`, borderRadius: 3, padding: '1px 6px' }}>{l}</span>
                              ))
                            : <span style={{ color: C.muted }}>—</span>
                          }
                        </div>
                      </Td>
                      <Td><span style={{ color: C.sub, fontWeight: 600 }} title={firstPhone(c.phones)}>{firstPhone(c.phones)}</span></Td>
                      <Td><span style={{ color: C.sub, fontWeight: 600 }} title={firstEmail(c.emails)}>{firstEmail(c.emails)}</span></Td>
                      <Td muted><span title={c.address}>{c.address || '—'}</span></Td>
                      <ActionCell onEdit={() => setCompanyModal({ ...c })} onDel={() => deleteCompany(c.id)} />
                    </tr>
                  );
                })
              }
            </tbody>
          </table>
        )}

        {/* Row count footer */}
        <div style={{ padding: '8px 14px', borderTop: `1px solid ${C.border}`, fontFamily: FF, fontSize: 11, color: C.muted, background: C.thBg, borderRadius: '0 0 8px 8px' }}>
          {tab === 'persons'
            ? `${visPersons.length} of ${persons.length} person${persons.length !== 1 ? 's' : ''}`
            : `${visCompanies.length} of ${companies.length} compan${companies.length !== 1 ? 'ies' : 'y'}`
          }
        </div>
      </div>

      {/* ── Modals */}
      {personModal  && <PersonModal  initial={personModal}  companies={companies} onSave={savePerson}  onClose={() => setPersonModal(null)}  />}
      {companyModal && <CompanyModal initial={companyModal}                        onSave={saveCompany} onClose={() => setCompanyModal(null)} />}
      {importModal   && <CrmImportModal onClose={() => setImportModal(false)} onImport={handleImport} />}
    </div>
  );
};

export default CrmBoard;
