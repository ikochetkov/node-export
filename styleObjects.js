/*
  Objects for styling cells - more about sythax - https://github.com/natergj/excel4node
*/

const header = {
  fill: {
    type: 'pattern',
    patternType: 'solid',
    bgColor: '#333333',
    fgColor: '#333333',
  },
  font: {
    size: 11,
    bold: true,
    color: 'white',
  },
  alignment: {
    horizontal: 'center',
    vertical: 'center'
  },
  border: {
    left: {
      style: 'thin',
      color: 'black'
    },
    right: {
      style: 'thin',
      color: 'black'
    },
    top: {
      style: 'thin',
      color: 'black'
    },
    bottom: {
      style: 'thin',
      color: 'black'
    }
  },
}

const regular = {
  font: {
    size: 11,
    bold: false,
    color: 'black',
  },
  alignment: {
    horizontal: 'right',
    vertical: 'center'
  },
  border: {
    left: {
      style: 'thin',
      color: 'black'
    },
    right: {
      style: 'thin',
      color: 'black'
    },
    top: {
      style: 'thin',
      color: 'black'
    },
    bottom: {
      style: 'thin',
      color: 'black'
    }
  },
}

const regular_highlighted = {
  fill: {
    type: 'pattern',
    patternType: 'solid',
    bgColor: '#d2d2d2',
    fgColor: '#d2d2d2',
  },
  font: {
    size: 11,
    bold: true,
    color: 'black',
  },
  alignment: {
    horizontal: 'right',
    vertical: 'center'
  },
  border: {
    left: {
      style: 'thin',
      color: 'black'
    },
    right: {
      style: 'thin',
      color: 'black'
    },
    top: {
      style: 'thin',
      color: 'black'
    },
    bottom: {
      style: 'thin',
      color: 'black'
    }
  },
}

module.exports = {
  header,
  regular,
  regular_highlighted
}