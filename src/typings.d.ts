// SCSS module type declarations for SPFx
declare module '*.module.scss' {
  const styles: { [className: string]: string };
  export default styles;
}

declare module '*.scss' {
  const styles: { [className: string]: string };
  export default styles;
}

// Asset type declarations
declare module '*.png' {
  const src: string;
  export default src;
}

declare module '*.ttf' {
  const src: string;
  export default src;
}
