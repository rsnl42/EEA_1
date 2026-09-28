module.exports = {
  content: ["./generate_dashboard.py", "./index.html"],
  theme: {
    extend: {
      fontFamily: {
        sans: ['"Plus Jakarta Sans"', 'sans-serif'],
      },
      colors: {
        brand: {
          50: '#eef2ff',
          100: '#e0e7ff',
          500: '#0072B2',
          600: '#005a8f',
          700: '#00436c',
          900: '#002b47',
        }
      }
    }
  }
}
