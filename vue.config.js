module.exports = {
  publicPath: './',
  assetsDir: 'static',
  productionSourceMap: false,
  devServer: {
    port: 8889, // 端口
    open: true,
    overlay: {
      warnings: false,
      errors: true
    }
  },
  configureWebpack: {
    // xlsx-style需要依赖于cptable，但是这个很大而且只有特殊情况才会使用，所以我们可以在打包的时候排除他
    externals: {
      './cptable': 'var cptable'
    }
  }
}
