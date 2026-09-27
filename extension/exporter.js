import Storage from './storage.js';

/**
 * Class Exporter
 */
export default class Exporter {
    constructor(userId = parseInt(location.search.substr(1))) {
        this.userId = parseInt(userId);
        this.workbook = XLSX.utils.book_new();
    }

    async exportInterest(storage) {
        let sheetNames = {
            "movie/done": "看过",
            "movie/doing": "在看",
            "movie/mark": "想看",
            "music/done": "听过",
            "music/doing": "在听",
            "music/mark": "想听",
            "book/done": "读过",
            "book/doing": "在读",
            "book/mark": "想读",
            "game/done": "玩过",
            "game/doing": "在玩",
            "game/mark": "想玩",
            "drama/done": "看过的舞台剧",
            "drama/mark": "想看的舞台剧",
            };
        for (let type of ["movie", "music", "book", "game", "drama"]) {
          for (let status of ["done", "doing", "mark"]) {
            let sheetName = sheetNames[`${type}/${status}`];
            if (!sheetName) continue;

            let collection = storage.local.interest
              .where({ type: type, status: status })
              .reverse();
            let data = [
              [
                "标题",
                "简介",
                "豆瓣评分",
                "链接",
                "创建时间",
                "我的评分",
                "标签",
                "评论",
                "可见性",
                "分类",
                "海报",
                "演员",
                "导演",
                "作者",
                "出版/上映日期",
                "出版社",
                "页数",
                "类型",
              ],
            ];
            await collection.each((row) => {
              let { subject, tags, rating, comment, create_time, is_private } =
                row.interest;
              console.log("[Nx] 导出数据", row);
              const intro = subject.intro ?? subject.card_subtitle ?? "";
              const press = subject.press ? subject.press[0] : "";
              const pages = subject.pages ? subject.pages[0] : "";
              data.push([
                subject.title,
                intro,
                subject.rating
                  ? subject.rating.value.toFixed(1)
                  : subject.null_rating_reason,
                subject.url,
                create_time,
                rating ? rating.value : "",
                tags.toString(),
                comment,
                is_private ? "private" : "public",
                subject.type,
                subject.pic.normal,
                JSON.stringify(subject.actors) ?? "",
                JSON.stringify(subject.directors) ?? "",
                JSON.stringify(subject.author) ?? "",
                JSON.stringify(subject.pubdate) ?? "",
                press,
                pages,
                JSON.stringify(subject.genres) ?? "",
              ]);
            });
            let worksheet = XLSX.utils.aoa_to_sheet(data);
            XLSX.utils.book_append_sheet(this.workbook, worksheet, sheetName);
          }
        }
      }

    async exportReview(storage) {
        let sheetNames = {'movie': '影评', 'music': '乐评', 'book': '书评', 'drama': '剧评', 'game': '游戏评论&攻略'};
        for (let type in sheetNames) {
            let collection = storage.local.review
                .where({ type: type })
                .reverse();
            let data = [['标题', '评论对象', '链接', '创建时间', '我的评分', '类型', '内容']];
            await collection.each(row => {
                let {
                    subject,
                    url,
                    rating,
                    fulltext,
                    title,
                    create_time,
                    type_name
                } = row.review;
                data.push([
                    title,
                    `《${subject.title}》`,
                    url,
                    create_time,
                    rating ? rating.value : '',
                    type_name,
                    fulltext,
                ]);
            });
            let worksheet = XLSX.utils.aoa_to_sheet(data);
            XLSX.utils.book_append_sheet(this.workbook, worksheet, sheetNames[type]);
        }
    }

    async exportAnnotation(storage) {
        let collection = storage.local.annotation
            .reverse();
        let data = [['书名', '章节', '页码', '链接', '创建时间', '我的评分', '内容']];
        await collection.each(row => {
            let {
                subject,
                chapter,
                page,
                url,
                rating,
                fulltext,
                create_time
            } = row.annotation;
            data.push([
                subject ? subject.title : '',
                chapter,
                page,
                url,
                create_time,
                rating ? rating.value : '',
                fulltext,
            ]);
        });
        let worksheet = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(this.workbook, worksheet, '笔记');
    }

    async exportStatus(storage) {
        let formatStatus = (status) => {
            if (status.deleted || status.hidden) {
                return status.msg;
            }
            let text = `${status.author.name}(@${status.author.uid})`;
            if (status.activity) {
                text += ` ${status.activity}`;
            }
            text += `: ${status.text}`;
            if (status.card) {
                text += `[推荐]:《${status.card.title}》(${status.card.url})`;
            }
            if (status.images && status.images.length > 0) {
                let images = [];
                status.images.forEach(image => {
                    images.push(image.large.url);
                });
                text += ` ${images}`;
            }
            if (status.parent_status) {
                text += `//${formatStatus(status.parent_status)}...`;
            }
            if (status.reshared_status) {
                text += `//${formatStatus(status.reshared_status)}`;
            }
            return text;
        };

        let collection = await storage.local.status
            .orderBy('id')
            .reverse();
        let data = [['创建时间', '链接', '内容', '话题']];
        await collection.each(row => {
            let {
                sharing_url,
                create_time,
                topic,
            } = row.status;
            data.push([
                create_time,
                sharing_url,
                formatStatus(row.status),
                topic ? [topic.title, topic.url].toString() : '',
            ]);
        });
        let worksheet = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(this.workbook, worksheet, '广播');
    }

    async exportFollowing(storage) {
        let data = [['用户名', '用户ID', '链接', '所在地', '备注']];

        let versionInfo = await storage.local.table('version').get({
            table: 'following',
        });

        if (versionInfo) {
            let collection = storage.local.following.where({ version: versionInfo.version });
            await collection.each(row => {
                let {
                    name,
                    uid,
                    url,
                    loc,
                    remark
                } = row.user;
                data.push([
                    name,
                    uid,
                    url,
                    loc ? loc.name : '',
                    remark,
                ]);
            });
        }

        let worksheet = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(this.workbook, worksheet, '我关注的');
    }

    async exportFollower(storage) {
        let data = [['用户名', '用户ID', '链接', '所在地']];

        let versionInfo = await storage.local.table('version').get({
            table: 'follower',
        });

        if (versionInfo) {
            let collection = storage.local.follower.where({ version: versionInfo.version });
            await collection.each(row => {
                let {
                    name,
                    uid,
                    url,
                    loc
                } = row.user;
                data.push([
                    name,
                    uid,
                    url,
                    loc ? loc.name : '',
                ]);
            });
        }
        let worksheet = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(this.workbook, worksheet, '关注我的');
    }

    async exportBlacklist(storage) {
        let data = [['用户名', '用户ID', '链接']];

        let versionInfo = await storage.local.table('version').get({
            table: 'blacklist',
        });

        if (versionInfo) {
            let collection = storage.local.blacklist.where({ version: versionInfo.version });
            await collection.each(row => {
                let {
                    name,
                    uid,
                    url
                } = row.user;
                data.push([
                    name,
                    uid,
                    url
                ]);
            });
        }
        let worksheet = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(this.workbook, worksheet, '黑名单');
    }

    async exportNote(storage) {
        let collection = storage.local.note.reverse();
        let data = [['标题', '链接', '创建时间', '修改时间', '内容']];
        await collection.each(row => {
            let {
                title,
                url,
                fulltext,
                create_time,
                update_time
            } = row.note;
            data.push([
                title,
                url,
                create_time,
                update_time,
                fulltext,
            ]);
        });
        let worksheet = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(this.workbook, worksheet, '日记');
    }

    async exportPhoto(storage) {
        let data = [['相册名称', '相册链接', '相册描述', '相册创建时间', '照片描述', '照片链接']];
        let albums = await storage.local.album.toArray();
        for (let {id, album} of albums) {
            data.push([album.title, album.url, album.description, album.create_time]);
            let photos = storage.local.photo.where({album: id});
            await photos.each(photo => {
                let {url, description} = photo.photo;
                data.push([null, null, null, null, description, url]);
            });
        }
        let worksheet = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(this.workbook, worksheet, '相册');
    }

    async exportDoumail(storage) {
        let data = [['用户', '链接', '发件人', '发送时间', '正文']];
        let contacts = await storage.local.doumailContact
            .orderBy('rank')
            .reverse()
            .toArray();
        for (let {id, contact, url} of contacts) {
            data.push([
                contact.name,
                url,
            ]);
            let doumails = storage.local.doumail.where({contact: id});
            await doumails.each(doumail => {
                let {content, sender, datetime} = doumail;
                data.push([null, null, sender.name, datetime, content]);
            });
        }
        let worksheet = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(this.workbook, worksheet, '豆邮');
    }

    async exportDoulist(storage) {
        let sheetNames = {'owned': '创建的豆列', 'following': '收藏的豆列'};
        for (let type in sheetNames) {
            let data = [['豆列名称', '豆列链接', '豆列简介', '豆列创建时间', '豆列更新时间', '内容名称', '内容链接', '来源', '评语']];
            let doulists = await storage.local.doulist.where({type: type}).toArray();
            for (let {id, doulist} of doulists) {
                data.push([
                    doulist.title,
                    doulist.url,
                    doulist.desc,
                    doulist.create_time,
                    doulist.update_time,
                ]);
                let items = storage.local.doulistItem.where({doulist: id});
                await items.each(item => {
                    let {url, title, source, comment} = item.item;
                    data.push([null, null, null, null, null, title, url, source, comment]);
                });
            }
            let worksheet = XLSX.utils.aoa_to_sheet(data);
            XLSX.utils.book_append_sheet(this.workbook, worksheet, sheetNames[type]);
        }
    }

    async exportBoard(storage) {
        let data = [['留言用户', '用户主页', '留言时间', '消息']];
        let messages = await storage.local.board
            .reverse()
            .toArray();
        for (let {id, sender, sendTime, message} of messages) {
            data.push([
                sender.name,
                sender.url,
                sendTime,
                message
            ]);
        }
        let worksheet = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(this.workbook, worksheet, '留言板');
    }

    async export(items) {
        let storage = new Storage(this.userId);
        await storage.local.open();
        try {
            for (let item of items) {
                switch (item) {
                case 'Interest':
                    await this.exportInterest(storage);
                    break;
                case 'Review':
                    await this.exportReview(storage);
                    break;
                case 'Annotation':
                    await this.exportAnnotation(storage);
                    break;
                case 'Status':
                    await this.exportStatus(storage);
                    break;
                case 'Following':
                    await this.exportFollowing(storage);
                    break;
                case 'Follower':
                    await this.exportFollower(storage);
                    break;
                case 'Blacklist':
                    await this.exportBlacklist(storage);
                    break;
                case 'Note':
                    await this.exportNote(storage);
                    break;
                case 'Photo':
                    await this.exportPhoto(storage);
                    break;
                case 'Doumail':
                    await this.exportDoumail(storage);
                    break;
                case 'Doulist':
                    await this.exportDoulist(storage);
                    break;
                case 'Board':
                    await this.exportBoard(storage);
                    break;
                }
            }
        } finally {
            storage.local.close();
        }
    }

    save() {
        let filename = `豆伴(${this.userId}).xlsx`;
        XLSX.writeFile(this.workbook, filename);
    }
}
