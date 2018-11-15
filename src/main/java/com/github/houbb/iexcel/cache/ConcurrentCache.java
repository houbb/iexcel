package com.github.houbb.iexcel.cache;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @author binbin.hou
 * @date 2018/11/15 17:30
 */
public class ConcurrentCache<K, V> implements ICache<K,V>{

    /**
     * 池
     */
    private final Map<K, V> cache = new ConcurrentHashMap<>();

    /**
     * 从缓存池中查找值
     *
     * @param key 键
     * @return 值
     */
    @Override
    public V get(K key) {
        return cache.get(key);
    }

    /**
     * 放入缓存
     *
     * @param key   键
     * @param value 值
     * @return 值
     */
    @Override
    public V put(K key, V value) {
        cache.put(key, value);
        return value;
    }

    /**
     * 移除缓存
     *
     * @param key 键
     * @return 移除的值
     */
    @Override
    public V remove(K key) {
        return cache.remove(key);
    }

    /**
     * 清空缓存池
     */
    @Override
    public void clear() {
        this.cache.clear();
    }

}
