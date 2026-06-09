import * as React from 'react';
import type { CrmAttachment } from './crmTypes';
import { fileToCrmAttachment } from './crmAttachmentStore';

const FF = 'Montserrat,sans-serif';

const isImageName = (name: string): boolean => /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(name);

const ml: React.CSSProperties = {
  fontSize: 10, fontWeight: 700, color: '#4a5568', letterSpacing: '.07em',
  textTransform: 'uppercase', marginBottom: 4, display: 'block', fontFamily: FF,
};

export const DocumentUploadSection: React.FC<{
  attachments: CrmAttachment[];
  onChange: (attachments: CrmAttachment[]) => void;
  readOnly?: boolean;
}> = ({ attachments, onChange, readOnly }) => {
  const inputRef = React.useRef<HTMLInputElement>(null);

  const addFiles = async (files: FileList | null): Promise<void> => {
    if (!files?.length) return;
    const added = await Promise.all(Array.from(files).map(fileToCrmAttachment));
    onChange([...attachments, ...added]);
  };

  const remove = (id: string): void => {
    onChange(attachments.filter(a => a.id !== id));
  };

  return (
    <div style={{ marginTop: 14 }}>
      <label style={ml}>Upload documents</label>
      {!readOnly && (
        <>
          <input
            ref={inputRef}
            type="file"
            multiple
            style={{ display: 'none' }}
            onChange={e => { void addFiles(e.target.files); e.target.value = ''; }}
          />
          <button
            type="button"
            onClick={() => inputRef.current?.click()}
            style={{
              padding: '8px 12px', background: '#f0f2f6', border: '1px dashed #cdd1d9',
              borderRadius: 4, color: '#8a97a8', fontFamily: FF, fontSize: 12,
              width: '100%', textAlign: 'left', cursor: 'pointer', boxSizing: 'border-box',
            }}
          >
            + Click to upload documents…
          </button>
        </>
      )}
      {attachments.length === 0 && readOnly && (
        <p style={{ fontFamily: FF, fontSize: 12, color: '#8a97a8', margin: 0 }}>No documents.</p>
      )}
      {attachments.length > 0 && (
        <div style={{ marginTop: 10, display: 'flex', flexWrap: 'wrap', gap: 10 }}>
          {attachments.map(att => {
            const isImg = att.dataUrl.startsWith('data:image') || isImageName(att.name);
            return (
              <div
                key={att.id}
                style={{
                  position: 'relative', border: '1px solid rgba(42,158,42,.35)', borderRadius: 4,
                  background: 'rgba(42,158,42,.09)', padding: 4, width: isImg ? 96 : 'auto', maxWidth: 220,
                }}
              >
                {isImg ? (
                  <img src={att.dataUrl} alt={att.name} style={{ width: 88, height: 88, objectFit: 'cover', borderRadius: 2, display: 'block' }} />
                ) : (
                  <a
                    href={att.dataUrl}
                    download={att.name}
                    style={{ display: 'block', padding: '6px 10px', fontSize: 11.5, color: '#2a9e2a', fontFamily: FF, textDecoration: 'none' }}
                  >
                    {att.name}
                  </a>
                )}
                {!readOnly && (
                  <button
                    type="button"
                    onClick={() => remove(att.id)}
                    title="Remove"
                    style={{
                      position: 'absolute', top: 2, right: 2, background: 'rgba(192,57,43,.9)', border: 'none',
                      color: '#fff', borderRadius: 2, width: 18, height: 18, fontSize: 11, cursor: 'pointer', lineHeight: 1,
                    }}
                  >
                    ×
                  </button>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};
