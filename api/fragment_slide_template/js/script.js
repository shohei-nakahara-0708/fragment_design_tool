(async () => {

  /** スワイプ閾値 */
  const SWIPE_THRESHOLD = 100;

  /**
   * Veeva APIリクエストの実行
   */
  function runAPIRequest(request) {
    console.log('runAPIRequest:', request);
    window?.webkit?.messageHandlers?.veeva.postMessage({'message': request});
  }

  /**
   * 次のスライドへ移動
   */
  function nextSlide() {
    console.log('nextSlide');
    runAPIRequest('veeva:nextSlide()');
  }

  /**
   * 前のスライドへ移動
   */
  function prevSlide() {
    console.log('prevSlide');
    runAPIRequest('veeva:prevSlide()');
  }

  /**
   * スワイプイベントのセットアップ
   */
  function setupSwipeEventsForBody() {
    let startX = 0;
    document.body.addEventListener('pointerdown', (event) => (startX = event.clientX));
    document.body.addEventListener('pointerup', (event) => {
      const deltaX = event.clientX - startX;
      if (Math.abs(deltaX) > SWIPE_THRESHOLD) deltaX > 0 ? prevSlide() : nextSlide();
    });
  }

  /**
   * アプリの初期化
   */
  let initializeApp;
  try {
    // shared.js から initializeApp をインポート
    const shared = await import('../shared/js/shared.js');
    initializeApp = shared.initializeApp;
  } catch (e) {
    // shared.js のインポートに失敗した場合、ローカルで初期化を行う
    initializeApp = () => setupSwipeEventsForBody();
  }
  // アプリの初期化を実行
  initializeApp();

})();
