import fs from 'fs';
import csvParser from 'csv-parser';
import { createObjectCsvWriter } from 'csv-writer';
import axios from 'axios';
import puppeteer from 'puppeteer';
import ExcelJS from 'exceljs';

const API_KEY = 'AIzaSyAkuKEZRVQD9b7VYfHeJfZOE466M7BvjNs'; // Replace with your actual API key

interface ChannelInfo {
  url: string;
  channelName: string;
  subscriberCount: string;
  dateJoined: string;
  videoCount: string;
  viewCount: string;
  location: string;
  emailUrl: string;
  instagramUrl: string;
  tiktokUrl: string;
  twitterUrl: string;
  facebookUrl: string;
  lastPost: string;
  channelStatus: string;
}

async function readCsv(filePath: string): Promise<string[]> {
  return new Promise((resolve, reject) => {
    const results: string[] = [];
    fs.createReadStream(filePath)
      .pipe(csvParser())
      .on('data', (data) => results.push(data.URL))
      .on('end', () => resolve(results))
      .on('error', (err) => reject(err));
  });
}

function extractHandleFromUrl(url: string): string | null {
  const regex = /youtube\.com\/@([a-zA-Z0-9_\-]+)(?:[\/?].*)?/;
  const match = url.match(regex);
  return match ? match[1] : null;
}

async function resolveChannelIdFromHandle(handle: string): Promise<string | null> {
  const url = `https://youtube.googleapis.com/youtube/v3/search?part=snippet&type=channel&q=${handle}&key=${API_KEY}`;
  try {
    const response = await axios.get(url);
    return response.data.items?.[0]?.id?.channelId || null;
  } catch (error) {
    console.error(`❌ Error resolving channel ID from handle: ${handle}`, error);
    return null;
  }
}

async function getLastVideoDate(playlistId: string): Promise<string> {
  const url = `https://youtube.googleapis.com/youtube/v3/playlistItems?part=snippet&maxResults=1&playlistId=${playlistId}&key=${API_KEY}`;
  try {
    const response = await axios.get(url);
    const date = response.data.items?.[0]?.snippet?.publishedAt;
    return date ? new Date(date).toLocaleDateString() : 'No posts yet';
  } catch (error) {
    return 'Error fetching date';
  }
}

async function getChannelInfo(channelId: string): Promise<ChannelInfo | null> {
  const url = `https://youtube.googleapis.com/youtube/v3/channels?part=snippet,statistics,contentDetails&id=${channelId}&key=${API_KEY}`;
  try {
    const response = await axios.get(url);
    const item = response.data.items?.[0];
    if (!item) return null;

    const uploadsPlaylistId = item.contentDetails?.relatedPlaylists?.uploads;
    const lastPost = uploadsPlaylistId
      ? await getLastVideoDate(uploadsPlaylistId)
      : 'No uploads';

    const lastPostYear = new Date(lastPost).getFullYear();
    const channelStatus = lastPostYear === 2025 ? 'Active' : 'Not Active';

    return {
      url: item.snippet.customUrl
        ? `https://www.youtube.com/${item.snippet.customUrl}`
        : `https://www.youtube.com/channel/${channelId}`,
      channelName: item.snippet.title,
      subscriberCount: item.statistics.subscriberCount,
      dateJoined: new Date(item.snippet.publishedAt).toLocaleDateString(),
      videoCount: item.statistics.videoCount,
      viewCount: item.statistics.viewCount,
      location: item.snippet.country || '',
      emailUrl: '',
      instagramUrl: '',
      facebookUrl: '',
      tiktokUrl: '',
      twitterUrl: '',
      lastPost,
      channelStatus
    };
  } catch (error) {
    console.error(`❌ Error fetching channel info for ID: ${channelId}`, error);
    return null;
  }
}

async function scrapeSocialLinks(channelUrl: string): Promise<Partial<ChannelInfo>> {
  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();
  try {
    await page.goto(`${channelUrl}/about`, { waitUntil: 'networkidle2', timeout: 60000 });
    const links = await page.evaluate(() => {
      const anchors = Array.from(document.querySelectorAll('a'));
      const getLink = (platform: string) =>
        anchors.find(a => a.href.includes(platform))?.href || '';
      return {
        facebookUrl: getLink('facebook.com'),
        instagramUrl: getLink('instagram.com'),
        tiktokUrl: getLink('tiktok.com'),
        twitterUrl: getLink('twitter.com')
      };
    });
    return links;
  } catch (err: any) {
    console.warn(`⚠️ Failed to scrape links from ${channelUrl}:`, err.message);
    return {};
  } finally {
    await browser.close();
  }
}

async function saveToCsv(data: ChannelInfo[], outputPath: string) {
  const csvWriter = createObjectCsvWriter({
    path: outputPath,
    header: [
      { id: 'url', title: 'URL' },
      { id: 'channelName', title: 'CHANNEL NAME' },
      { id: 'subscriberCount', title: 'TOTAL SUBSCRIBERS' },
      { id: 'dateJoined', title: 'DATE JOINED' },
      { id: 'videoCount', title: 'TOTAL VIDEOS' },
      { id: 'viewCount', title: 'TOTAL VIEWS' },
      { id: 'location', title: 'LOCATION' },
      { id: 'emailUrl', title: 'EMAIL URL' },
      { id: 'instagramUrl', title: 'INSTAGRAM URL' },
      { id: 'tiktokUrl', title: 'TIKTOK URL' },
      { id: 'twitterUrl', title: 'X URL' },
      { id: 'facebookUrl', title: 'FB URL' },
      { id: 'lastPost', title: 'LAST POST' },
      { id: 'channelStatus', title: 'CHANNEL STATUS' }
    ]
  });
  await csvWriter.writeRecords(data);
  console.log('✅ CSV saved to output/results.csv');
}

async function saveToExcel(data: ChannelInfo[], outputPath: string) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('YouTube Data');

  const headerStyle = {
    font: { bold: true },
    alignment: { vertical: 'middle', horizontal: 'center' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDDDDD' } }
  };

  const cellStyle = {
    alignment: { vertical: 'middle', horizontal: 'left' }
  };

  const alternatingRowStyle = {
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } }
  };

  sheet.columns = [
    { header: 'URL', key: 'url', width: 40 },
    { header: 'CHANNEL NAME', key: 'channelName', width: 30 },
    { header: 'TOTAL SUBSCRIBERS', key: 'subscriberCount', width: 20 },
    { header: 'DATE JOINED', key: 'dateJoined', width: 20 },
    { header: 'TOTAL VIDEOS', key: 'videoCount', width: 15 },
    { header: 'TOTAL VIEWS', key: 'viewCount', width: 20 },
    { header: 'LOCATION', key: 'location', width: 20 },
    { header: 'EMAIL URL', key: 'emailUrl', width: 30 },
    { header: 'INSTAGRAM URL', key: 'instagramUrl', width: 30 },
    { header: 'TIKTOK URL', key: 'tiktokUrl', width: 30 },
    { header: 'X URL', key: 'twitterUrl', width: 30 },
    { header: 'FB URL', key: 'facebookUrl', width: 30 },
    { header: 'LAST POST', key: 'lastPost', width: 20 },
    { header: 'CHANNEL STATUS', key: 'channelStatus', width: 20 }
  ];

  data.forEach((info, index) => {
    const row = sheet.addRow(info);
    const isNotActive = info.channelStatus === 'Not Active';
  
    row.eachCell((cell) => {
      Object.assign(cell.style, cellStyle);
  
      // Apply alternating row background if needed
      if (index % 2 === 1 && !isNotActive) {
        Object.assign(cell.style, alternatingRowStyle);
      }
  
      // 🔴 If status is "Not Active", color the whole row red
      if (isNotActive) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFC7CE' } // Light red
        };
      }
    });
  });
  

  sheet.getRow(1).eachCell((cell) => Object.assign(cell.style, headerStyle));
  await workbook.xlsx.writeFile(outputPath);
  console.log(`📥 Excel export complete: ${outputPath}`);
}

async function main() {
  try {
    console.log('📥 Reading CSV from input/link.csv...');
    const urls = await readCsv('input/link.csv');
    console.log(`🔍 Found ${urls.length} URLs to process.`);

    const results: ChannelInfo[] = [];

    for (const [index, rawUrl] of urls.entries()) {
      const url = rawUrl.trim();
      if (!url) continue;

      console.log(`\n🔗 [${index + 1}/${urls.length}] Processing: ${url}`);

      const handle = extractHandleFromUrl(url);
      if (!handle) {
        console.warn(`⚠️ Skipping: Couldn't extract handle from URL.`);
        continue;
      }
      console.log(`   👉 Extracted handle: @${handle}`);

      const channelId = await resolveChannelIdFromHandle(handle);
      if (!channelId) {
        console.warn(`⚠️ Skipping: Failed to resolve channel ID.`);
        continue;
      }
      console.log(`   ✅ Resolved channel ID: ${channelId}`);

      const info = await getChannelInfo(channelId);
      if (!info) {
        console.warn(`⚠️ Skipping: Couldn't fetch channel info.`);
        continue;
      }
      console.log(`   📺 Channel: ${info.channelName}`);
      console.log(`   📅 Last Post: ${info.lastPost} → ${info.channelStatus}`);

      const socialLinks = await scrapeSocialLinks(info.url);
      Object.assign(info, socialLinks);
      results.push(info);
    }

    console.log('\n💾 Saving data to CSV...');
    await saveToCsv(results, 'output/results.csv');

    console.log('📊 Saving data to Excel...');
    await saveToExcel(results, 'output/results.xlsx');

    console.log('\n✅ All done!');
  } catch (error) {
    console.error('❌ Error in main():', error);
  }
}


main();
