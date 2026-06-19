/**
 * 네이버 블로그 이미지 다운로더 전용 CORS 프록시 (Cloudflare Worker)
 *
 * 사용법: https://<당신의-워커>.workers.dev/?url=<인코딩된 대상 URL>
 *
 * - 대상 URL의 HTML/이미지를 그대로 받아 CORS 헤더를 붙여 돌려줍니다.
 * - 네이버 차단을 피하려고 브라우저처럼 User-Agent / Referer 를 붙입니다.
 * - 무료 요금제: 하루 10만 요청 (개인 도구에는 충분)
 */
export default {
  async fetch(request) {
    const reqUrl = new URL(request.url);
    const target = reqUrl.searchParams.get('url');

    // CORS preflight 처리
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders() });
    }

    if (!target) {
      return new Response('Missing ?url= parameter', {
        status: 400,
        headers: corsHeaders(),
      });
    }

    let targetUrl;
    try {
      targetUrl = new URL(target);
    } catch {
      return new Response('Invalid url', { status: 400, headers: corsHeaders() });
    }

    // 안전장치: http/https 만 허용
    if (targetUrl.protocol !== 'http:' && targetUrl.protocol !== 'https:') {
      return new Response('Only http/https allowed', { status: 400, headers: corsHeaders() });
    }

    try {
      const upstream = await fetch(targetUrl.toString(), {
        headers: {
          'User-Agent':
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' +
            '(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
          'Accept': 'text/html,application/xhtml+xml,image/avif,image/webp,*/*',
          'Accept-Language': 'ko-KR,ko;q=0.9,en;q=0.8',
          // 네이버는 Referer 가 자기 도메인일 때 차단을 덜 함
          'Referer': 'https://blog.naver.com/',
        },
        redirect: 'follow',
      });

      // 원본 응답을 그대로 스트리밍하면서 CORS 헤더만 덧붙임
      const resp = new Response(upstream.body, upstream);
      const headers = corsHeaders();
      for (const [k, v] of Object.entries(headers)) {
        resp.headers.set(k, v);
      }
      // 캐싱 (같은 이미지/글 반복 호출 줄이기)
      resp.headers.set('Cache-Control', 'public, max-age=3600');
      return resp;
    } catch (e) {
      return new Response('Proxy fetch failed: ' + e.message, {
        status: 502,
        headers: corsHeaders(),
      });
    }
  },
};

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': '*',
  };
}
