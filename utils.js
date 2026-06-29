const sleep = (ms) => new Promise(r => setTimeout(r, ms));

const randomDelay = (min = 3000, max = 7000) => {
  const ms = Math.floor(Math.random() * (max - min + 1)) + min;
  return sleep(ms);
};

const getFirstName = (fullName) => {
  return (fullName || "").trim().split(" ")[0];
};

module.exports = { sleep, randomDelay, getFirstName };
